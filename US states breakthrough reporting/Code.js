// Constants
const WORKSHEET_SHEET_ID = 206055979;
const MATRIX_SHEET_ID = 123679669;

// Fragile constants, update these whenever we add a new state or make changes to the metadata we include here
// States that have data
const FIRST_STATE = "AK";
const LAST_STATE = "WV";

// States Matrix sheet row numbers 
const MATRIX_SHEET_HEADER_ROW_NUMBER = 1;
const MATRIX_SHEET_LINKS_ROW_NUMBER = 2;
const MATRIX_SHEET_NOTES_ROW_NUMBER = 3;
const MATRIX_SHEET_SCREENSHOTS_ROW_NUMBER = 5;

// States Matrix sheet row numbers 
const WQRKSHEET_SHEET_HEADER_ROW_NUMBER = 1;
const WQRKSHEET_SHEET_SCREENSHOTS_ROW_NUMBER = 3;

// Reinitialize with user confirmation 
function confirmReinitializeWorksheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Reinitializing Worksheet Settings', 'This will clear all notes in the Worksheet input columns, and reinitialize them based on the States Matrix. It will also clear and reset the Screenshot checkboxes. Continue?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    if (!reinitializeSourceNotes() || (!reinitializeScreenshotCheckboxes) || (!reinitializeSourceLinks)) {
      ui.alert("There was an error.");
    }
  } else {
    ui.alert("This action has been cancelled.");
  }
}

// sets the values of the source links in worksheet based on states matrix tab 
function reinitializeFormatting() {
  const sheet = getSheetById(WORKSHEET_SHEET_ID);
  const firstCol = findHeaderColumn(WORKSHEET_SHEET_ID, `${WQRKSHEET_SHEET_HEADER_ROW_NUMBER}:${WQRKSHEET_SHEET_HEADER_ROW_NUMBER}`,FIRST_STATE);
  const lastCol = findHeaderColumn(WORKSHEET_SHEET_ID, `${WQRKSHEET_SHEET_HEADER_ROW_NUMBER}:${WQRKSHEET_SHEET_HEADER_ROW_NUMBER}`,LAST_STATE); 
  const range = sheet.getRange(3,4,45,2) ;//lastCol-firstCol+1);
  range.setNumberFormat('#,##0');
  for (let row; row<46;row++){
    
  }
  Logger.log(range.getA1Notation());
  return true;
}

// sets the values of the source links in worksheet based on states matrix tab 
function reinitializeSourceLinks() {
  const matrixSourceLinksRange = getOneRowRange(MATRIX_SHEET_ID,
    MATRIX_SHEET_LINKS_ROW_NUMBER,
    MATRIX_SHEET_HEADER_ROW_NUMBER);  
  const  matrixSheetName = getSheetById(MATRIX_SHEET_ID).getName();
  const sheet = getSheetById(WORKSHEET_SHEET_ID);
  const firstCol = findHeaderColumn(WORKSHEET_SHEET_ID, `${WQRKSHEET_SHEET_HEADER_ROW_NUMBER}:${WQRKSHEET_SHEET_HEADER_ROW_NUMBER}`,FIRST_STATE);  
  sheet.getRange(MATRIX_SHEET_LINKS_ROW_NUMBER,firstCol).setFormula(`=ARRAYFORMULA(HYPERLINK('${matrixSheetName}'!${matrixSourceLinksRange.getA1Notation()},"↗"))`);
  return true;
}

// sets the values of the screenshots checkboxes based on states matrix tab
function reinitializeScreenshotCheckboxes() {
  const worksheetScreenshotCheckboxesRange = getOneRowRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_SCREENSHOTS_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  const matrixScreenshotCheckboxesRange = getOneRowRange(MATRIX_SHEET_ID,
    MATRIX_SHEET_SCREENSHOTS_ROW_NUMBER,
    MATRIX_SHEET_HEADER_ROW_NUMBER);  
    initializeWorksheetCheckboxes(worksheetScreenshotCheckboxesRange, matrixScreenshotCheckboxesRange);
    return true;
}

// sets the values of notes based on states matrix tab
function reinitializeSourceNotes() {
  const worksheetInputRange = getOneRowRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  const matrixNotesRange = getOneRowRange(MATRIX_SHEET_ID,
    MATRIX_SHEET_NOTES_ROW_NUMBER,
    MATRIX_SHEET_HEADER_ROW_NUMBER); 

  const matrixCols = matrixNotesRange.getNumColumns();
  const matrixRows = matrixNotesRange.getNumRows();
  const matrixNotes = matrixNotesRange.getValues();     
  
  for(let r = 0; r<matrixRows; r++){
    for(let c = 0; c<matrixCols; c++) {
      if(matrixNotes[r][c]) {
        matrixNotes[r][c] = matrixNotes[r][c] // + "\n\n\n======\nDon't edit these notes, they will be overwritten."
      }
    }
  } 

  initializeWorksheetComments(worksheetInputRange, matrixNotesRange);
  return true;
}

function clearRangeNotes(range) {
  range.clearNote();
}

function getOneRowRange(sheetId, rowId, headerRowId) {
    const sheet = getSheetById(sheetId);
    const firstCol = findHeaderColumn(sheetId, `${headerRowId}:${headerRowId}`,FIRST_STATE);
    const lastCol = findHeaderColumn(sheetId, `${headerRowId}:${headerRowId}`,LAST_STATE);    
    return sheet.getRange(rowId, firstCol, 1, lastCol-firstCol+1);
}

function initializeWorksheetComments(worksheetInputRange, matrixNotesRange) {
  Logger.log(worksheetInputRange.getA1Notation());
  Logger.log(matrixNotesRange.getA1Notation());
  
  const matrixCols = matrixNotesRange.getNumColumns();
  const matrixRows = matrixNotesRange.getNumRows();
  const matrixNotes = matrixNotesRange.getValues();

  Logger.log(`matrixCols = ${matrixCols}\t matrixRows = ${matrixRows}\t matrixNotes = ${matrixNotes}\t`);


  for(let r = 0; r<matrixRows; r++){
    for(let c = 0; c<matrixCols; c++) {
      if(matrixNotes[r][c]) {
        matrixNotes[r][c] = matrixNotes[r][c] // + "\n\n\n======\nDon't edit these notes, they will be overwritten."
      }
    }
  } 
        
  worksheetInputRange.setNotes(matrixNotes);
}

function initializeWorksheetCheckboxes(worksheetInputRange, matrixCheckboxesRange) {
  Logger.log(`initializeWorksheetCheckboxes\n worksheetInputRange= ${worksheetInputRange.getA1Notation()}`);
  Logger.log(`${matrixCheckboxesRange.getA1Notation()}`);
  
  const matrixCols = matrixCheckboxesRange.getNumColumns();
  const matrixRows = matrixCheckboxesRange.getNumRows();
  const matrixNotes = matrixCheckboxesRange.getValues();

  Logger.log(`matrixCols = ${matrixCols}\t matrixRows = ${matrixRows}\t matrixNotes = ${matrixNotes}\t`);
  worksheetInputRange.insertCheckboxes();

  for(let r = 0; r<matrixRows; r++){
    for(let c = 0; c<matrixCols; c++) {
      if(matrixNotes[r][c]) {
        matrixNotes[r][c] = true ;
      } else {
        matrixNotes[r][c] = false;
      }
    }
  } 
        
  worksheetInputRange.setValues(matrixNotes);
}

function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}

function findHeaderColumn(sheetId, rowSelector, label){
  const sheet = getSheetById(sheetId);
  const data = sheet.getRange(rowSelector).getValues();
  for(let i = 0; i<data[0].length; i++){
    if(data[0][i] == label) {
      return i+1;
    }
  }
}



