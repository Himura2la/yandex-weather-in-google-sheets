var dataRows = 100
var APIKey = "" // https://developer.tech.yandex.ru
var icons = {
  'day': '🏙️',
  'evening': '🌆',
  'night': '🌃',
  'clear': '☀️',
  'clear-night': '🌙',
  'overcast': '☁️',
  'partly-cloudy':  '🌤️',
  'cloudy': '⛅',
  'overcast-thunderstorms-with-rain': '⛈️',
  'overcast-and-light-rain': '🌦️',
  'cloudy-and-rain': '🌧️',
  'overcast-and-rain': '☔',
  'partly-cloudy-and-light-rain': '🌦️',
  'partly-cloudy-and-rain': '🌧️',
  'cloudy-and-light-rain': '🌧️'
}

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
  var resp = requestAPI("https://api.weather.yandex.ru/v1/forecast?lang=ru_RU&l10n=true&lat=" + lat + "&lon=" + lon)
  var l10n = resp['l10n']

  function iconize(part, condition){
    if (part == 'night') {
      switch(condition) {
        case 'clear':
          condition += '-night'
          break;
      }
    }
    if (condition in icons)
      return icons[part] + " " + icons[condition]
    return icons[part] + " " + l10n[condition]
  }
  
  var days = {}
  for each (var f in resp['forecasts']) {
    var day_c = ''
    for each (var part in ['day', 'evening', 'night']) {
      var c = f['parts'][part]
      day_c += iconize(part, c['condition']) + ' ' + c['temp_min'] + '..' + c['temp_max'] + "    "
    }
    // day_c += '🌇 ' + f['sunset']
    days[f['date']] = day_c
  }
  return days
}


function updateWeather() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getActiveSheet();
  var timeZone = spreadsheet.getSpreadsheetTimeZone()
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var dayCol = findIndex(header, "Day")
  var weatherCol = findIndex(header, "Yandex Weather")
  
  var lastUpdate = "🔄 " + Utilities.formatDate(new Date(), timeZone, "dd.MM HH:mm")
  sheet.getRange(2, weatherCol).setValue(lastUpdate)
  
  var forecast = getForecast(54.2, 37.6)
  
  for (var row = 2; row < dataRows; row++) {
    var day = sheet.getRange(row, dayCol).getValue()
    if (!day) continue
    
    var key = Utilities.formatDate(day, timeZone, "yyyy-MM-dd")
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
