// function testEventBasedFunction(){
//   var total = getExpenseRatioYahoo("QQQ");
//   Logger.log(total)
// }

// getExpenseRatio.gs
//  v0.04 - 12/3/2021 - baseline cleanup for sharing.
//  v0.05a - 12/3/2021 - fixed forceRefresh and retrieve from cache.

/**
 * Get expense ratio from yahoo or cache for a fund ticker.
 * @param {"VTI"} ticker Fund ticker symbol.
 * @param {false} forceRefresh [OPTIONAL, default = FALSE]. true forces a new quote, not use any available cache.
 * @returns expenseRatio
 * @customfunction
 */
function getCachedExpenseRatioYahoo(ticker, forceRefresh = false) {

  var cacheKey = "ExpenseRatioYahoo-" + ticker;
  var cache = CacheService.getPublicCache();

  // retrieve a previously cached ER if it has not timed out or forced quote by caller.
  if (!forceRefresh) {
    var cacheVal = cache.get(cacheKey);
    if (cacheVal != null) {
      if (cacheVal == "" || cacheVal == "N/A" || cacheVal == "n/a")
        return "n/a";
      return parseFloat(cacheVal);
    }
  }

  resp = getExpenseRatioYahoo(ticker);
  // set ttl to 30 days, ER doesn't change often.
  var ttl = 30 * 24 * 60 * 60;
  logCacheTTL(cacheKey, "", ttl);
  cache.put(cacheKey, resp, ttl);
  return resp;
}

/**
 * Get expense ratio for fund ticker from yahoo.
 * Returns market price for ETFs, and NAV for mutual funds.
 * @param {"VTI"} ticker Fund ticker symbol.
 * @param {false} includeChangePct [OPTIONAL, default = FALSE]. 1 or true = return change % column; 0, false or missing = do not.
 * @returns [asOf, price] or [asOf, price, change %].
 * @customfunction
 */
function getExpenseRatioYahoo(ticker) {
  // Fetch HTML from URL.
  // tag to search for: EXPENSE_RATIO-value
  // https://finance.yahoo.com/quote/VTSAX?p=VTSAX&.tsrc=fin-srch

  var url = "https://finance.yahoo.com/quote/" + ticker + "/?p=" + ticker + "&.tsrc=fin-srch";
  Logger.log(`Url: ${url}`)

  var resp = UrlFetchApp.fetch(url).getContentText();
  //console.log(`Response ${resp}`)
  var stringER = _Yahoo_extractExpenseRatio(resp);
  Logger.log(`Expense ratio: ${stringER}`)
  if (stringER == "" || stringER == "N/A")
    return "n/a";
  return stringER / 100;
}

function _Yahoo_extractExpenseRatio(inputMarkup) {
  inputMarkup = inputMarkup.substring(inputMarkup.indexOf("EXPENSE_RATIO-value"));
  inputMarkup = inputMarkup.substring(inputMarkup.indexOf(">"));
  inputMarkup = inputMarkup.substring(1, inputMarkup.indexOf("</td>"));
  inputMarkup = inputMarkup.replace(/">/, "");
  inputMarkup = inputMarkup.replace(/%/, "");
  return inputMarkup;
}
