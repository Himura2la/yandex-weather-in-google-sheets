var dataRows = 100
var APIKey = "" // https://developer.tech.yandex.ru

var API = "https://api.weather.yandex.ru/v1"

function requestAPI(url) {
  try {    
    var response = UrlFetchApp.fetch(url, {headers: {"X-Yandex-API-Key": APIKey}})
    return JSON.parse(response.getContentText())
  } catch (e) {
    Logger.log(e)
  }
  return "UrlFetchApp failed."
}


function getForecast(lat, lon){
  var resp = requestAPI(API + "/forecast?lang=ru_RU&l10n=true&lat=" + lat + "&lon=" + lon)
  var days = {}
  var l = resp['l10n']

  function ic(s){
    switch (s) {
      case 'day':
        return '⛺'
      case 'evening':
        return '🌆'
      case 'night':
        return '🌃'
      case 'clear':
        return '☀️'
      case 'overcast':
        return '☁️'
      case 'partly-cloudy':
        return '🌤️'
      case 'cloudy':
        return '⛅'
      case 'overcast-thunderstorms-with-rain':
        return '⛈️'
      case 'overcast-and-light-rain':
        return '🌦️'
      case 'cloudy-and-rain':
        return '🌧️'
      case 'overcast-and-rain':
        return '☔'
      case 'partly-cloudy-and-light-rain':
        return '🌦️'
      case 'partly-cloudy-and-rain':
        return '🌧️'
      case 'cloudy-and-light-rain':
        return '🌧️'
    }
    return l[s]
  }
  
  for each (var f in resp['forecasts']) {
    var day_c = ''
    for each (var part in ['day', 'evening', 'night']) {
      var c = f['parts'][part]
      day_c += ic(part) + " " + ic(c['condition']) + ' ' + c['temp_min'] + '..' + c['temp_max'] + "\t"
    }
    day_c += '🌇 ' + f['sunset']
    days[f['date']] = day_c
  }
  return days
}


function updateWeather() {
  var forecast = getForecast(55.7, 37.6)
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var dayCol = findIndex(header, "Day")
  var weatherCol = findIndex(header, "Yandex Weather")
  
  for (var row = 2; row < dataRows; row++) {
    var day = sheet.getRange(row, dayCol).getValue()
    if (day == "") continue
    
    var key = Utilities.formatDate(day, "GMT", "yyyy-MM-dd")
    var f = forecast[key]
    if (f != undefined)
      sheet.getRange(row, weatherCol).setValue(f)
  }
}

function findIndex(array, item) {
  for (var i=0; i<array.length; i++)
    if (array[i] == item) return i + 1;
  return -1;
} 

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.addMenu("Yandex Weather", [{name: "Обновить погоду",functionName: "updateWeather"}]);
}
