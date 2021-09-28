function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Pandemic Tracking')

  .addItem("--- Pivot Table", "empty")
  .addItem('          Recreate pivot table', 'createPivotTable')
  .addToUi();
}
