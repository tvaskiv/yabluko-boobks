function myFunction() {
  var url = "https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?json&valcode=usd";
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText();
  var data = JSON.parse(json);
  Logger.log(data[0].rate);
}
