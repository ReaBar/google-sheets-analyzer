/**
* Returns etfs data.
*
* @param {"VOO"} tickers - Input ticker or tickers.
* @param {true} headers - Add or remove headers.
* @return {array} etfs data.
* @customfunction
*/
function ETFS(tickers, headers = true) {
  console.log(tickers);
  if (!Array.isArray(tickers)) {
    tickers = [tickers]
  };

  const valuesToFormat = ["closingPrice", "closingPriceChng", "closingPriceChngPct", "volume"];
  const output = [];

  const requests = tickers.flat().map(tic => {
    return {
      url: 'https://real-time.etf.com/',
      method: 'post',
      headers: {
        Authorization: 'Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCIsImtpZCI6Ik1qSTBOa1pDTTBaR1JUTTROemRFUkRVM1JrRXdOelUxTkRoRVFrUTRNakpEUlRsQk9FWXdSUSJ9.eyJpc3MiOiJodHRwczovL2V0ZmRvdGNvbS5hdXRoMC5jb20vIiwic3ViIjoiNmVVNnVFWEYydVpKY25uVVZlMFJxRlNuR3lWbEM1NkRAY2xpZW50cyIsImF1ZCI6Imh0dHBzOi8vcmVhbC10aW1lLmV0Zi5jb20iLCJpYXQiOjE2Mzk5MzMxODQsImV4cCI6MTY0MDAxOTU4NCwiYXpwIjoiNmVVNnVFWEYydVpKY25uVVZlMFJxRlNuR3lWbEM1NkQiLCJndHkiOiJjbGllbnQtY3JlZGVudGlhbHMifQ.AutORBdl9597gITotTLltnb3QUpGRY0iQLd3SKM3FBNlsMH-wMihl2deVcA6tqyR-uRf_hXyOzIRQiIc5qBnICU4Hxaqq-wTc4qiaG5iZLGQLzi6rFk64xU0dBy0O6C8La4Y9xQiFZA5PODdqSgUBBRgDL5IDvodlY1IqBUSzNSfjXFPLxGlCXAo08wwG8ESoIGA7NDAzdJ_fCk-yMtYUC0d3X60QTBPBF5xp43K6qdM6GytFkR6KMou-ASxmzRVX6kdnc4G_CSFeAELi9zzMgP6IDJXEVmdnk0-379WPOASW1rJ35-4DSvdZuxyyB0YocJCu7-PkKAam7ygwO_xwA',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        "operationName": "QuoteByTicker",
        "variables": {
          "ticker": tic
        },
        "query": "query QuoteByTicker($ticker: String!) {  quoteByTicker(ticker: $ticker) {    ticker    marketSession    closingPrice    closingPriceDate    closingPriceTime    closingPriceChng    closingPriceChngPct    formattedClosingPrice    formattedClosingPriceDate    formattedClosingPriceChngPct    lastTrade    lastTradeDate    lastTradeTime    change    changePct    formattedLastTrade    formattedLastTradeDate    formattedChangePct    volume    attribution    bid    bidSize    ask    askSize    __typename  }}"
      })
    }
  });

  const responses = UrlFetchApp.fetchAll(requests);
  responses.forEach((res, i) => {
    const data = JSON.parse(res.getContentText()).data.quoteByTicker;
    if (i == 0 && headers) {
      output.push(Object.keys(data));
    }
    const tempArray = [];

    Object.entries(data).forEach(entrie => {
      const [key, value] = entrie;
      if (valuesToFormat.includes(key)) {
        tempArray.push(Number(value))
      } else {
        tempArray.push(value)
      }
    })

    output.push(tempArray);

  })

  return output;

}

