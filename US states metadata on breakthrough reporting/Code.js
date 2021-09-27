// To change the lookups, use this
const SPLIT_TO_COLOUMNS = "C2";
const STATE_IMPORT = "C1"
const ROW_COUNTER = "A1"
const ROW_HEADERS = "A2"
const AIRTABLE_SHEET = "Airtable";
const RESULTS_SHEET = "Results";
let stateColumns = [];
const FIRST_ROW = 2;
const CHECK_VALUE = "X";
const COUNT_COLUMN = "B";
const METRIC_COLORS = {
  "BI cases": "#3d85c6",
  "BI deaths": "#073763",
  "BI hosp": "#0b5394",
  "Fully": "#274e13",
  "not BI cases": "#b4a7d6",
  "not BI death": "#674ea7",
  "not BI hosp": "#8e7cc3",
  "Total": "#f1c232"
};
const FIRST_AIRTABLE_SPLIT_COLUMN="C";
const LAST_AIRTABLE_SPLIT_COLUMN="Z";
const rowNames = [];

function NUM_RETURN_LETRA(column){
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


function createPivotTable() {
  Logger.log(`createPivotTable`);
  // Airtable tab, column C: split comma delimited values into rows 
  let range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AIRTABLE_SHEET).getRange(SPLIT_TO_COLOUMNS);
  range.setFormula(`=ArrayFormula(SPLIT(B2:B35,","))`);

  // Results tab: import states from airtable sheet & display as columns
  range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET).getRange(STATE_IMPORT);
  range.setFormula(`=TRANSPOSE(${AIRTABLE_SHEET}!A2:A)`); 
 

  // create the row headers
  Logger.log(`creating row headers`);
  const airtableLastRow =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AIRTABLE_SHEET).getLastRow();
  const airtableLastColumn = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AIRTABLE_SHEET).getLastColumn();
  console.log(`Airtable: last row = ${airtableLastRow} \n last Column = ${airtableLastColumn}`);
  range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET).getRange(ROW_HEADERS);
  range.setFormula(`SORT(UNIQUE(TRANSPOSE(SPLIT(ARRAYFORMULA(CONCATENATE(TRANSPOSE(FILTER(${AIRTABLE_SHEET}!C2:${NUM_RETURN_LETRA(airtableLastColumn)}${airtableLastRow},${AIRTABLE_SHEET}!C2:C${airtableLastRow}>0)&"-"))),"-"))))`);  

  // count the number of rows (unique values reported by states)
    Logger.log(`counting number of rows`);
  const countRange  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET).getRange(ROW_COUNTER);
  countRange.setFormula(`=COUNTA(${ROW_HEADERS}:A)`);
  const lastRow = Number(countRange.getValue())+1 ;
  console.log(`lastRow=  ${lastRow}`);

  // figure out what state columns there are, starting with column C (3) 
  const RESULTS_STATE_GRID_START_COLOUMN = 3
  for( i = RESULTS_STATE_GRID_START_COLOUMN; i< (airtableLastRow-1)+RESULTS_STATE_GRID_START_COLOUMN; i++)
    stateColumns.push(NUM_RETURN_LETRA(i));     

  // popoulate the first data column containing counts for each metric
  Logger.log(`populating count column`);
  for (let number=FIRST_ROW;number<=lastRow;number++) {
    const cellAddress = `${COUNT_COLUMN}${number}`;     
    const cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET).getRange(cellAddress);
    cell.setFormula(`=COUNTIF(${stateColumns[0]}${number}:${stateColumns[stateColumns.length-1]}${number},"X")`);
  }

    Logger.log(`setting the colors of the first column based on the metric name`);
  // set the colors of each metric in the first column based on the metric name
  for (let number=FIRST_ROW;number<=lastRow;number++) {
    const cellAddress = `A${number}`; 
    const cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET).getRange(cellAddress);
    cell.setFontColor("#ffffff");
    cell.setFontWeight("bold");
    const cellValue = cell.getValue();
    rowNames.push(cellValue);
    cell.setHorizontalAlignment("left");
    for(let propt in METRIC_COLORS){
        if (cellValue.search(propt)==0) {
          cell.setBackground(METRIC_COLORS[propt]);
        }        
    }
  }

  // set the values of each cell in the matrix, as well as statewise counts under each state column
  Logger.log(`setting the values of each cell as well as statewise counts`);
  for(let letter = 0; letter < stateColumns.length; letter++) {
    const rowInAirtableSheet = letter+2;
    const cellAddress = `${stateColumns[letter]}${lastRow+1}`; 
    const cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET).getRange(cellAddress);
    cell.setFormula(`=COUNTIF(${stateColumns[letter]}1:${stateColumns[letter]}${lastRow},"X")`);    

    for (let number=FIRST_ROW;number<=lastRow;number++) {
      const cellAddress = `${stateColumns[letter]}${number}`;
      const cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET).getRange(cellAddress);
      cell.setFormula(`=IF(IFNA(MATCH($A$${number},${AIRTABLE_SHEET}!${FIRST_AIRTABLE_SPLIT_COLUMN}${rowInAirtableSheet}:${LAST_AIRTABLE_SPLIT_COLUMN}${rowInAirtableSheet},0),"")="","","X")`);
    }
  }

Logger.log(rowNames);

 Logger.log(`setting the format of each row of state data as percent or number based on the row name`);
  for (let number=FIRST_ROW;number<=lastRow;number++) {
      const rangeAddress = `${stateColumns[0]}${number}:${stateColumns[stateColumns.length-1]}${number}`;
      Logger.log(`number = ${number}\t rangeAddress = ${rangeAddress}\t`);
      const range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTS_SHEET).getRange(rangeAddress);
      range.clearFormat();
      range.setHorizontalAlignment("center");
      // figure out the name of the row
      const firstCellValue = rowNames[number-2];
      Logger.log(`firstCellValue = ${firstCellValue}\t firstCellValue.search("percent")=${firstCellValue.search("percent")}`);   
      if ((firstCellValue.search("percent")>0) || (firstCellValue.search(" per ")>0)) {
            range.setNumberFormat("#0.##0");
            Logger.log(`setting ${rangeAddress} format to percent`);
          } else {
              range.setNumberFormat("#,##0");
          }    
  }          
}  


  


/**
 * Returns the Hexadecimal value of a cell's background color.
 *
 * @param {number} row The cell's row number.
 * @param {number} column The cell's column number.
 * @return The Hexadecimal value of the cell's background color.
 * @customfunction
 */
function BGHEX(row, column) {
  var background = SpreadsheetApp.getActive().getDataRange().getCell(row, column).getBackground();
  return background;
}


