function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Pandemic Tracking')

  .addItem('          Reinitialize Worksheet', 'confirmReinitializeWorksheet')
  .addSeparator()
  .addItem('          Clear Checkers, DCers and last check', 'confirmReinitializeChecks')
  .addItem('          Reinitialize Notes', 'confirmReinitializeNotes')
  .addItem('          Reinitialize Screenshot Status', 'confirmReinitializeScreenshotsStatus')
  .addSeparator()
  .addItem('          Create QA Sheet', 'createQASheet')

  .addToUi();
}
