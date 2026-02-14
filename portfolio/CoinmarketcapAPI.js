function test_getCryptoPrice(){
  var symbol = "ADA"
  getCryptoPrice(symbol)
}

function getCryptoPrice(symbol) {
  var sh2=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet2");
  
  //Make sure that you got the API key from Coinmarketcap API dashboard and paste it in sheet_1 on cell B1
  var apiKey="e52b99df-687a-4166-a4c3-7fd03896b42e"
  
  var url=`https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol=${symbol}`
  var requestOptions = {
  method: 'GET',
  uri: 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest',
  qs: {
    start: 1,
    limit: 5000,
    convert: 'USD'
  },
  headers: {
    'X-CMC_PRO_API_KEY': apiKey
  },
  json: true,
  gzip: true
};
  
  var httpRequest= UrlFetchApp.fetch(url, requestOptions);
  var getContext= httpRequest.getContentText();
  
  var parseData=JSON.parse(getContext);
  Logger.log(`parsed info: ${parseData.data[symbol].quote.USD.price}`);
  return parseData.data[symbol].quote.USD.price;
}