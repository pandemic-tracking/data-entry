function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Pandemic Tracking')

  .addItem("--- Configuration", "empty")
  .addItem('          Recreate Worksheet Settings', 'confirmReinitializeWorksheet')
  .addToUi();
}
