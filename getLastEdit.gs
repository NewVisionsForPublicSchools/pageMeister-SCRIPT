function pageMeister_getLastEdit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var data = getRowsData(sheet);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastUpdateIndex = headers.indexOf("Last Updated")+1;
  for (var i=0; i<data.length; i++) {
    var url = data[i].linkToPage;
    if (url) {
      try {
        var page = SitesApp.getPageByUrl(url);
        var children = page.getAllDescendants();
        var allTimes = [];
        allTimes.push(page.getLastUpdated().getTime());
        for (var j=0; j<children.length; j++) {
          allTimes.push(children[j].getLastUpdated().getTime());
        }
        var lastUpdate = Math.max.apply(null,allTimes);
        lastUpdate = new Date(lastUpdate);
        lastUpdate = Utilities.formatDate(lastUpdate, ss.getSpreadsheetTimeZone(), "M/d/yy H:m a");
        sheet.getRange(i+2, lastUpdateIndex).setValue(lastUpdate);
      } catch (err) {
        sheet.getRange(i+2, lastUpdateIndex).setValue(err.message);
      }
    }
  }
}
