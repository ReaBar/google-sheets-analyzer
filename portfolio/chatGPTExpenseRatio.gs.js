function getMorningstarHtml(url) {
  var options = {
    method: "get",
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    },
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, options);
  return response.getContentText();
}
