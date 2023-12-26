const dateRange = ["A4:A23","E1"];

function setDate() {
  var range = SpreadsheetApp.getActive().getSheetByName("Caisse");
  var androany = new Date().toLocaleDateString('fr');
  for (n=0;n<dateRange.length;n++) {
    range.getRange(dateRange[n]).setValue(androany);
  }
}
