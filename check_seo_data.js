function countOutlinks() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outlinksSheet = spreadsheet.getSheetByName("Outlinks");
  var uniqueOutlinksSheet = spreadsheet.getSheetByName("Unique Outlinks");

  if (!outlinksSheet) {
    outlinksSheet = spreadsheet.insertSheet("Outlinks");
  }
  if (!uniqueOutlinksSheet) {
    uniqueOutlinksSheet = spreadsheet.insertSheet("Unique Outlinks");
  }

  var currentDate = new Date().toLocaleDateString("de-DE");
  var lastOutlinkColumn = outlinksSheet.getLastColumn() + 1;
  var lastUniqueOutlinkColumn = uniqueOutlinksSheet.getLastColumn() + 1;

  var urls = outlinksSheet.getRange(2, 1, outlinksSheet.getLastRow() - 1, 1).getValues();

  urls.forEach(function(url, index) {
    url = url[0];
    if (url) {
      try {
        var html = UrlFetchApp.fetch(url).getContentText();
        var outlinks = html.match(/<a [^>]*href="(\/[^\/"][^"]*|https:\/\/www\.aponeo\.de[^"]*)"/g);
        if (outlinks) {
          var uniqueOutlinks = [...new Set(outlinks)];

          outlinksSheet.getRange(1, lastOutlinkColumn).setValue(currentDate);
          outlinksSheet.getRange(index + 2, lastOutlinkColumn).setValue(outlinks.length);

          uniqueOutlinksSheet.getRange(1, lastUniqueOutlinkColumn).setValue(currentDate);
          uniqueOutlinksSheet.getRange(index + 2, lastUniqueOutlinkColumn).setValue(uniqueOutlinks.length);
        }
      } catch (error) {
        console.error("Fehler beim Abrufen der URL " + url + ": " + error.message);

        outlinksSheet.getRange(index + 2, lastOutlinkColumn).setValue("Fehler");
        uniqueOutlinksSheet.getRange(index + 2, lastUniqueOutlinkColumn).setValue("Fehler");
      }
    }
  });
}

function checkResponseTimes() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var responseTimesSheet = spreadsheet.getSheetByName("Response Times");

  if (!responseTimesSheet) {
    responseTimesSheet = spreadsheet.insertSheet("Response Times");
  }

  var currentDate = new Date().toLocaleDateString("de-DE");
  var lastResponseTimeColumn = responseTimesSheet.getLastColumn() + 1;

  var urls = responseTimesSheet.getRange(2, 1, responseTimesSheet.getLastRow() - 1, 1).getValues();

  urls.forEach(function(url, index) {
    url = url[0];
    if (url) {
      try {
        var startTime = new Date().getTime();
        var response = UrlFetchApp.fetch(url);
        var endTime = new Date().getTime();

        var responseTime = endTime - startTime;

        responseTimesSheet.getRange(1, lastResponseTimeColumn).setValue(currentDate);
        responseTimesSheet.getRange(index + 2, lastResponseTimeColumn).setValue(responseTime);
      } catch (error) {
        console.error("Fehler beim Abrufen der URL " + url + ": " + error.message);

        responseTimesSheet.getRange(index + 2, lastResponseTimeColumn).setValue("Fehler");
      }
    }
  });
}
