function onEdit(e) {
  var range = e.range.getSheet();
  var SheetName = range.getName()
  //Logger.log(SheetName);
  var LastRow = range.getLastRow();
  //Logger.log(LastRow);
  range.getRange(1,LastRow).setNumberFormat("mmm dd, yyyy")
}
