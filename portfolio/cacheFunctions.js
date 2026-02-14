// getExpenseRatio.gs
//  v0.04 - 12/3/2021 - baseline cleanup for sharing.
//  v0.05 - 12/8/2021 - added a ttl timeout calculator for instrumentation purposes only. Also continued to tweak the TTL algorithm.

/**
 * Get asOfDateTime, price, and optionally change % from chosen data source or its cache for a fund ticker.
 * Returns market price for ETFs, and NAV for mutual funds.
 * @param {yahoo} source available data sources include "bigcharts","morningstar", "vanguardFundID", "yahoo"
 * @param {true} includeChangePct [OPTIONAL, default = FALSE]. 1 or true = return change % column; 0, false or missing = do not.@param {"VTI"} ticker Fund ticker symbol.
 * @param {false} includeChangePct [OPTIONAL, default = FALSE]. 1 or true = return change % column; 0, false or missing = do not.
 * @param {false} forceRefresh [OPTIONAL, default = FALSE]. true forces a new quote, not use any available cache.
 * @returns [asOf, price] or [asOf, price, change %].
 * @customfunction
 */
function getCachedPrices(source = "bigcharts", ticker = "VTI", includeChangePct = false, forceRefresh = false) {

  // For now, I am short-cutting any calls to morningstar due to its 100% failure rate.
  switch (source) {
    case "TSP":
      // we are not retrieving change % for TSP, force false.
      includeChangePct = false;
      break;
    case "morningstar":
      throw source + " quotes busted, not allowing calls at this time.";
  }

  var cacheKey = source + "-" + ticker;
  var cache = CacheService.getPublicCache();

  // retrieve a previously cached quote if it has not timed out or forced quote by caller.
  if (!forceRefresh) {
    var cacheVal = cache.get(cacheKey);
    // parse the stored JSON if it exists and return to the caller.
    if (cacheVal != null) {
      var cacheArray = JSON.parse(cacheVal);
      // convert JSON string date back to a javascript date. Man, dates are annoying.
      if (cacheArray[0][0])
        cacheArray[0][0] = new Date(cacheArray[0][0]);

      if (includeChangePct) {
        return cacheArray;
      } else {
        return [[cacheArray[0][0], cacheArray[0][1]]];
      }
    }
  }

  var resp;
  // Using a switch to call these functions as I don't know the security implications of the run function.
  switch (source) {
    case 'bigcharts':
      resp = getPriceBigCharts(ticker, true);
      break;
    case 'morningstar':
      resp = getPriceMstar(ticker, true);
      break;
    case 'TSP':
      resp = getPriceTSP(ticker);
      break;
    case 'vanguardFundID':
      resp = getVanguardPriceFundID(ticker);
      break;
    case 'yahoo':
      resp = getPriceYahoo(ticker, true);
      break;
    default:
      throw ("invalid data source '" + source + "'");
  }

  // log TTL calculations.
  var tradetime = resp[0][0]
  var ttl = calcCacheTTL(tradetime);
  logCacheTTL(cacheKey, tradetime, ttl);

  cache.put(cacheKey, JSON.stringify(resp), ttl);
  if (includeChangePct) {
    return resp;
  } else {
    return [[resp[0][0], resp[0][1]]];
  }
}

/**
 * Calculate an age on the quote and use that to estimate how long it may remain useful in the cache (in seconds)..
 * @param {"11-26-2021 13:59:00"} t tradetime used for quote aging calculations.
 * @customfunction
 */
function calcCacheTTL(t) {

  // get the time as it is known for the markets.
  var nycDate = _getNYCTime();

  // get the time and day of the most recent trade of the reference ticker symbol, time passed in as a parameter.
  // if t is unset, we use midnight yesterday as a default.
  var tradetime;
  if (t) {
    tradetime = new Date(Date.parse(t));
  } else {
    tradetime = new Date;
    tradetime.setDate(tradetime.getDate() - 1);
    tradetime.setHours(0);
    tradetime.setMinutes(0);
  }
  var tradetimeDayOfMonth = tradetime.getDate();
  var tradetimeMonth = tradetime.getMonth();
  var tradetimeYear = tradetime.getFullYear();

  // If last trade is within an hour of current time, we assume market is open.
  if (Math.abs(tradetime.getTime() - nycDate.getTime()) < (1000 * 60 * 60)) {
    return 20 * 60;
  }

  // breakdown NYC time.
  var nycMinutes = nycDate.getMinutes();
  var nycHours = nycDate.getHours();
  var nycDayOfWeek = nycDate.getDay();   // Day of Week.
  var nycDayOfMonth = nycDate.getDate(); // The day of the month.
  var nycMonth = nycDate.getMonth();
  var nycYear = nycDate.getFullYear();

  // If it is the weekend, let's cache for 9 hours.
  // This will force refreshes on Monday morning at 9 at the latest.
  if (nycDayOfWeek == 0 || nycDayOfWeek == 6)
    return (nycHours < 9) ? (9 - nycHours) * 60 * 60 : (24 + 9 - nycHours) * 60 * 60;

  // If there time is other than 00:00, it is actively updated during the day.
  // Note, this still lets most mutual fund queries go through, but I haven't figured out a way to identify all of them.
  if (tradetime.getHours()) {

    // if the trade date is at least 14 days ago, no need to work hard during the trading day.
    if ((nycDate - tradetime) > (1000 * 60 * 60 * 24 * 14)) {
      // If time is before 3pm, update expected between 6:45 and 8pm today.
      // Update every 30 minutes between 7pm and 9pm, other hours, evenings, go 4 just in case.
      if (nycHours < 18)
        return (18 - nycHours) * 60 * 60;
      if (nycHours == 18)
        return (60 - nycMinutes) * 60;
      if (nycHours > 18 && nycHours <= 20)
        return 30 * 60;
      return 4 * 60 * 60;
    }

    // If current time is between 7AM and 9AM, then the market could be in pre-open
    // ex-Dividend changes often show up then.
    if (nycHours >= 7 && nycHours < 9) {
      return 30 * 60;
    }

    // Calculate the remaining time before market open at 9:30
    // 10 minute grace period to allow to start.
    if (nycHours == 9 && nycMinutes < 40) {
      return (40 - nycMinutes) * 60;
    }

    // we are in trading hours and we have the stupid yahoo MF quote, let's just push it to 4pm.
    if (tradetimeDayOfMonth == nycDayOfMonth && tradetime.getHours() == 8) {
      // If time is before 5pm, update expected between 5:45 and 7pm today.
      // Update every 20 minutes between 6pm and 9pm, other hours, evenings, go 4 just in case.
      if (nycHours < 17)
        return (17 - nycHours) * 60 * 60;
      if (nycHours == 17)
        return (60 - nycMinutes) * 60;
      if (nycHours > 17 && nycHours <= 20)
        return 20 * 60;
      // Otherwise make it 4 hours.
      return 4 * 60 * 60;
    }

    // If we are in trading hours but no current trade, then 30 minutes to catch when it trades.
    if (nycHours < 16)
      return 30 * 60;

    // If current time is between 4pm and 7pm,
    // market could be in the process of closing the books for today.
    if (nycHours >= 16 && nycHours < 19) {
      // Let's calculate the lesser of 60 minutes or the remaining time before 7pm true up.
      return (nycHours == 18) ? (60 - nycMinutes) * 60 : 60 * 60;
    }

    // If tradetime is < today and current time is between 7pm and 9pm,
    // use 20 minutes to ensure we have the latest and final daily update.
    if ((tradetimeYear <= nycYear && tradetimeMonth <= nycMonth && tradetimeDayOfMonth < nycDayOfMonth) &&
      (nycHours >= 19 && nycHours < 22)) {
      return 20 * 60;
    }

    // Otherwise make it 3 hours.
    return 3 * 60 * 60;
  }

  // if we get here, it is a date only value, then set to update after 4pm only if the date is earlier than today.
  if (tradetimeYear <= nycYear && tradetimeMonth <= nycMonth && tradetimeDayOfMonth < nycDayOfMonth) {
    if (nycHours < 17)
      return (17 - nycHours) * 60 * 60;
    if (nycHours == 17)
      return (60 - nycMinutes) * 60;
    if (nycHours > 17 && nycHours <= 20)
      return 20 * 60;
  }
  // Fallback, make it 3 hours.
  return 3 * 60 * 60;
}

function _getNYCTime() {
  // current date in local timezone with seconds trimmed.
  var currDate = new Date();
  currDate.setSeconds(0);
  currDate.setMilliseconds(0);

  // This is considered unsafe, but it appears to work.
  // It basically gets a string representing the current time in NYC timezone regardless of
  // the current time zone defined in the user's environment and converts to a new Date thus showing the
  // local time as it is seen in NYC, not here, wherever here is.....
  var nycDateString = currDate.toLocaleString('en-US', { timeZone: 'America/New_York' });
  return new Date(Date.parse(nycDateString));
}

/************* These functions are used for instrumenting the the TTL calculations for analysis purposes. ***********/
// log TTL calculations to a worksheet for analysis of
// effectiveness. they are put in a cache themselves and I
// will pull them via a function call in the spreadsheet.
function logCacheTTL(tag, tradetime, ttl) {
  // there is overhead here, put in a return if you want insturmenting turned off (production)
  // these caching items actually seem to add significant overhead to the system and impact response times.
  return;
  var cache = CacheService.getPublicCache();
  var resp = [tradetime, _getNYCTime(), ttl];
  cache.put("tts-" + tag, JSON.stringify(resp), ttl + (10 * 60));
}

/**
 * Given a known cache tag, return the logged TTL and times if recorded.
 * @param {"bigcharts-QUOTE-VTI"} tag cache tag.
 * @returns [tradetime, nycTime, ttl]
 * @customfunction
 */
function getCacheTTL(tag = "vanguardSummary") {
  var cache = CacheService.getPublicCache();
  var cacheVal = cache.get("tts-" + tag);
  if (cacheVal) {
    var resp = JSON.parse(cacheVal);
    if (resp) {
      if (resp[0]) resp[0] = new Date(resp[0]);
      if (resp[1]) resp[1] = new Date(resp[1]);
      return [resp];
    }
  }
}

/*
 * Add seconds to the input startTime and return the new time the cache will timeout.
 * @param {"11-26-2021 13:59:00"} startTime current Time used to calcluate the original TTL.
 * @param {"11-26-2021 13:59:00"} ttlSecs TTL calculated originally.
 * @customfunction
 * */
function ttlTimeout(startTime, ttlSecs) {
  var t = new Date(startTime);
  t.setSeconds(t.getSeconds() + ttlSecs);
  return t;
}