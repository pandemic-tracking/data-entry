// Constants
const WORKSHEET_SHEET_ID = 206055979; 
const MATRIX_SHEET_ID = 291219688; /// 1090654428; 
const SOURCE_NOTES_SHEET_ID = 752396001; 
// Fragile constants, update these whenever we add a new state or make changes to the metadata we include here
// States that have data
const FIRST_STATE = "AK";
const LAST_STATE = "WY";

// States Matrix sheet row numbers 
const MATRIX_SHEET_HEADER_ROW_NUMBER = 1;
const MATRIX_SHEET_LINKS_ROW_NUMBER = 2;
const MATRIX_SHEET_PRIMARY_SCREENSHOTS_NOTES_ROW_NUMBER = 5;
const MATRIX_SHEET_SECONDARY_SCREENSHOTS_NOTES_ROW_NUMBER = 12;
const MATRIX_SHEET_NOTES_ROW_NUMBER = 14;
const MATRIX_SHEET_PRIMARY_SCREENSHOTS_STATUS_ROW_NUMBER = 15;
const MATRIX_SHEET_SECONDARY_SCREENSHOTS_STATUS_ROW_NUMBER = 16;


// Worksheet sheet row numbers 
const WQRKSHEET_SHEET_HEADER_ROW_NUMBER = 1;
const WQRKSHEET_SHEET_PRIMARY_URL_ROW_NUMBER = 2;
const WQRKSHEET_SHEET_PRIMARY_SCREENSHOTS_ROW_NUMBER = 4;
const WQRKSHEET_SHEET_SECONDARY_SCREENSHOTS_ROW_NUMBER = 5;
const WQRKSHEET_SHEET_DATA_START_ROW_NUMBER = 6;
const WQRKSHEET_SHEET_LASTCHECK_ROW_NUMBER = 52;
const WQRKSHEET_SHEET_CHECKER_ROW_NUMBER = 53;
const WQRKSHEET_SHEET_DCER_ROW_NUMBER = 54;

const SOURCE_NOTES_SHEET_HEADER_ROW_NUMBER = 1;
const SOURCE_NOTES_SHEET_FIRST_NOTE_ROW_NUMBER = 2;

// Reinitialize with user confirmation 
function confirmReinitializeWorksheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(`⚠️ Reinitializing Worksheet Settings`, `This will: 
  (1) Clear checker, double checker, and last checked  
  (2) Clear all popup notes, and reinitialize them based on States Matrix (airtable)
  (3) Reset screenshot status columns based on States Matrix (airtable)
  
  Continue?`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    if (!reinitializeChecks()){
      ui.alert("There was an error reinitializing checks");
    }
    if (!reinitializeSourceNotes()) {
      ui.alert("There was an error reinitializing source notes");
    }
    if (!reinitializeScreenshotsStatus()){
      ui.alert("There was an error reinitializing screenshot status");
    }
  } else {
    ui.alert("This action has been cancelled.");
  }
}

// Reinitialize checks with user confirmation 
function confirmReinitializeChecks() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(`Reinitializing Checks`, `This will: 
  clear checker, double checker, and last checked  
  Continue?`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    if (!reinitializeChecks()){
      ui.alert("There was an error.");
    }    
  } else {
    ui.alert("This action has been cancelled.");
  }
}

// sets the values of notes based on states matrix tab
function reinitializeChecks() {
  Logger.log(`\n********\nReinitializng Checks\n********`);
  let worksheetInputRange = getOneRowRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_LASTCHECK_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  worksheetInputRange.clearContent();

  worksheetInputRange = getOneRowRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_CHECKER_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  worksheetInputRange.clearContent();

  worksheetInputRange = getOneRowRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_DCER_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  worksheetInputRange.clearContent();
  return true;
}

// Reinitialize screenshot status with user confirmation 
function confirmReinitializeScreenshotsStatus() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(`Reinitializing Screenshot Status`, `This will: 
  Reset screenshot status columns based on States Matrix (from airtable)
  
  Continue?`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    if (!reinitializeScreenshotsStatus()){
      ui.alert("There was an error.");
    }
  } else {
    ui.alert("This action has been cancelled.");
  }
}

// sets the values of screenshot status based on states matrix tab
function reinitializeScreenshotsStatus() {
  Logger.log(`********
  Reinitializng primary screenshots status
  ********`);
  const worksheetInputRange = getOneRowRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_PRIMARY_SCREENSHOTS_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  const matrixNotesRange = getOneRowRange(MATRIX_SHEET_ID,
    MATRIX_SHEET_PRIMARY_SCREENSHOTS_STATUS_ROW_NUMBER,
    MATRIX_SHEET_HEADER_ROW_NUMBER); 
  worksheetInputRange.setValues(matrixNotesRange.getValues());

  Logger.log(`********
  Reinitializng secondary screenshots status
  ********`);
  const  wsSecondary = getOneRowRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_SECONDARY_SCREENSHOTS_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  const mnSecondary = getOneRowRange(MATRIX_SHEET_ID,
    MATRIX_SHEET_SECONDARY_SCREENSHOTS_STATUS_ROW_NUMBER,
    MATRIX_SHEET_HEADER_ROW_NUMBER); 
  wsSecondary.setValues(mnSecondary.getValues());
  return true;
}

// Reinitialize with user confirmation 
function confirmReinitializeNotes() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(`Reinitializing Notes`, `This will: 
  Clear all popup notes, and reinitialize them based on States Matrix (from airtable)
  
  Continue?`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    if (!reinitializeSourceNotes()) {
      ui.alert("There was an error.");
    }
  } else {
    ui.alert("This action has been cancelled.");
  }
}

// sets the values of notes based on states matrix tab
function reinitializeSourceNotes() {
  Logger.log(`
  Reinitializng Source Notes
  `);
  const worksheetInputRange = getOneRowRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  const matrixNotesRange = getOneRowRange(MATRIX_SHEET_ID,
    MATRIX_SHEET_NOTES_ROW_NUMBER,
    MATRIX_SHEET_HEADER_ROW_NUMBER); 
  initializeWorksheetNotes(worksheetInputRange, matrixNotesRange);

  Logger.log(`\n*****************************\nReinitializng Manual Screenshot Notes for primary screenshot URL`);
  const wsPrimaryScreenshotRange = getOneRowRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_PRIMARY_SCREENSHOTS_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  const matrixPrimaryScreenshotNotesRange = getOneRowRange(MATRIX_SHEET_ID,
    MATRIX_SHEET_PRIMARY_SCREENSHOTS_NOTES_ROW_NUMBER,
    MATRIX_SHEET_HEADER_ROW_NUMBER); 
  initializeWorksheetNotes(wsPrimaryScreenshotRange, matrixPrimaryScreenshotNotesRange);

  Logger.log(`\n*****************************\nReinitializng Manual Screenshot Notes for secondary screenshot URL`);
  const wsSecondaryScreenshotRange = getOneRowRange(WORKSHEET_SHEET_ID,
        WQRKSHEET_SHEET_SECONDARY_SCREENSHOTS_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER); 
  const matrixSecondaryScreenshotNotesRange =  getOneRowRange(MATRIX_SHEET_ID,
    MATRIX_SHEET_SECONDARY_SCREENSHOTS_NOTES_ROW_NUMBER,
    MATRIX_SHEET_HEADER_ROW_NUMBER); 
  initializeWorksheetNotes(wsSecondaryScreenshotRange, matrixSecondaryScreenshotNotesRange);

Logger.log(`
  Reinitializng Detailed Source Notes
  `);
  const ws2InputRange = getRange(WORKSHEET_SHEET_ID,
    WQRKSHEET_SHEET_DATA_START_ROW_NUMBER,
    WQRKSHEET_SHEET_HEADER_ROW_NUMBER,45); 
  const sourceNotesRange = getRange(SOURCE_NOTES_SHEET_ID,
    SOURCE_NOTES_SHEET_FIRST_NOTE_ROW_NUMBER,
    SOURCE_NOTES_SHEET_HEADER_ROW_NUMBER,
    45); 
  initializeWorksheetNotes(ws2InputRange, sourceNotesRange);


  return true;
}


function clearRangeNotes(range) {
  range.clearNote();
}

function getOneRowRange(sheetId, rowId, headerRowId) {
    Logger.log(`calling getOneRowRange for sheet ${sheetId} with rowId = ${rowId}  and headerRowId = ${headerRowId}`);
    const sheet = getSheetById(sheetId);
    const firstCol = findHeaderColumn(sheetId, `${headerRowId}:${headerRowId}`,FIRST_STATE);
    const lastCol = findHeaderColumn(sheetId, `${headerRowId}:${headerRowId}`,LAST_STATE); 
    Logger.log(`firstCol = ${firstCol}  and lastCol = ${lastCol}`);   
    return sheet.getRange(rowId, firstCol, 1, lastCol-firstCol+1);
}

function getRange(sheetId, rowId, headerRowId,numRows) {
    Logger.log(`calling getRange for sheet ${sheetId} with rowId = ${rowId}  and headerRowId = ${headerRowId}`);
    const sheet = getSheetById(sheetId);
    const firstCol = findHeaderColumn(sheetId, `${headerRowId}:${headerRowId}`,FIRST_STATE);
    const lastCol = findHeaderColumn(sheetId, `${headerRowId}:${headerRowId}`,LAST_STATE); 
    Logger.log(`firstCol = ${firstCol}  and lastCol = ${lastCol}`);   
    return sheet.getRange(rowId, firstCol, numRows, lastCol-firstCol+1);
}

function initializeWorksheetNotes(inputRange, notesRange) {
  Logger.log(`Worksheet: ${inputRange.getA1Notation()}`);
  Logger.log(`States Matrix: ${notesRange.getA1Notation()}`);
  
  const matrixCols = notesRange.getNumColumns();
  const matrixRows = notesRange.getNumRows();
  const matrixNotes = notesRange.getValues();

  Logger.log(`matrixCols = ${matrixCols}\t matrixRows = ${matrixRows}\t matrixNotes = ${matrixNotes}\t`);
  inputRange.setNotes(matrixNotes);
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
    //Logger.log(`i=${i}  data[0][i]=${data[0][i]} and label = ${label}`)
    if(data[0][i] == label) {
      return i+1;
    }
  }
}

function createQASheet() {
  const SHEET_URL = "https://docs.google.com/spreadsheets/d/1c8kXHBylT6dVqzW67PCIYqu31EI2OgzeEk9r9lU5QUs/edit#gid=0";
  const SCRIPT_URL = "http://54.172.243.72:8080/api/bi-checks";

  const htmlOutput = HtmlService
    .createHtmlOutput(`Go to <a href="${SCRIPT_URL}">this site</a> for spreadsheet checks`)
    .setWidth(250) //optional
    .setHeight(50); //optional

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(`This will create a new tab in BI spreadsheet checks sheet. 
  
  Continue?`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    try {      
      const options = { method: "get" }
      const res = UrlFetchApp.fetch(SCRIPT_URL, options);
      SpreadsheetApp.getActive().toast(htmlOutput,"🎉Success");
  } catch(err) {
      ui.alert(err);
      Logger.log(err);
    }
  }
}

