var props = PropertiesService.getUserProperties();
props.setProperty("actionPagePurchaseFirstRow", 2);
props.setProperty("weightedAvgFirstRow", "R5");
props.setProperty("valueHistoryFirstRow" ,48);
props.setProperty("cryptoActionsFirstRow", 21);
props.setProperty("cryptoActionsLastRow", 48);
props.setProperty("cryptoSummaryFirstRow",19);
props.setProperty("cryptoActionsPage", "קריפטו פעולות");
props.setProperty("summaryPage", "תמונת מצב");
props.setProperty("cleanWorth", "שווי נקי");

//var scriptProperties = PropertiesService.getScriptProperties();

function testEventBasedFunction(){
  var debug_e = {
    'day-of-month': 01,
    'month': 12,
    'year': 2024
  };
  var total = totalValueHistoryNewSheet(debug_e);
  Logger.log(total)
}

function totalValueHistoryFromButton() {
  totalValueHistoryNewSheet();
}

// DEPRACATED
function totalValueHistory(e){
  var currentMonth = e['month'];
  var currentYear = e['year'];

  var currentDate = Utilities.formatDate(new Date(currentYear, currentMonth), "GMT", "MM/yyyy");
  Logger.log("Current date: "+currentDate);
  var historyDateLastRow = findLastRow("תמונת מצב", Number(props.getProperty("valueHistoryFirstRow")), "K");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = spreadsheet.getSheetByName('תמונת מצב');
  Logger.log("summary history last row: "+historyDateLastRow);
  var totalValueWithoutPensionCell = spreadsheet.getRangeByName("worth_total_value_without_pension").getA1Notation();
  var totalValueWithoutPension = summarySheet.getRange(totalValueWithoutPensionCell).getValue();
  Logger.log("total value without pension: " + totalValueWithoutPension);
  summarySheet.getRange("K"+historyDateLastRow).setValue(currentDate);
  summarySheet.getRange("L"+historyDateLastRow).setValue(totalValueWithoutPension);
}

function totalValueHistoryNewSheet(e){
  var currentMonth
  var currentYear
  if (e && typeof e === 'object' && 'month' in e && 'year' in e){
    currentMonth = e['month'];
    currentYear = e['year'];
  } else {
    var now = new Date();
    currentMonth = now.getMonth() + 1;
    currentYear = now.getFullYear();
  }

  Logger.log("Month: " + currentMonth + ", Year: " + currentYear);
  // var currentMonth = e['month'];
  // var currentYear = e['year'];
  Logger.log("current month: "+currentMonth);
  var currentDate = Utilities.formatDate(new Date(currentYear, currentMonth), "GMT", "MM/yyyy");
  Logger.log("Current date: "+currentDate);
  var historyDateLastRow = findLastRow(props.getProperty("cleanWorth"), Number(1), "A") - 1;
  Logger.log("Clean worth sheet last row: "+historyDateLastRow);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var cleanWorthSheet = spreadsheet.getSheetByName(props.getProperty("cleanWorth"));
  var summarySheet = spreadsheet.getSheetByName(props.getProperty("summaryPage"));
  var totalValueWithoutPensionCell = spreadsheet.getRangeByName("worth_total_value_without_pension").getA1Notation();
  var totalValueWithoutPension = summarySheet.getRange(totalValueWithoutPensionCell).getValue();
  var relevantColumn = currentMonth+1;
  Logger.log("total value without pension: " + totalValueWithoutPension);
  Logger.log(`current month ${currentMonth},column to edit ${relevantColumn}, last row ${historyDateLastRow}`);
  cleanWorthSheet.getRange(historyDateLastRow, relevantColumn).setValue(totalValueWithoutPension);
}

function addYearRow() {
  var today = new Date();
  
  // Check if it's January 1st
  if (today.getMonth() === 0 && today.getDate() === 1) {
    var currentYear = today.getFullYear();
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var cleanWorthSheet = spreadsheet.getSheetByName(props.getProperty("cleanWorth"));
    var lastRow = cleanWorthSheet.getLastRow();
    
    // Insert a new row at the end of the sheet
    cleanWorthSheet.insertRowAfter(lastRow);
    
    // Set the year in column A
    cleanWorthSheet.getRange(lastRow + 1, 1).setValue(currentYear);
    
    // Copy formatting from the previous row
    var previousRowRange = cleanWorthSheet.getRange(lastRow, 1, 1, cleanWorthSheet.getLastColumn());
    var newRowRange = cleanWorthSheet.getRange(lastRow + 1, 1, 1, cleanWorthSheet.getLastColumn());
    previousRowRange.copyTo(newRowRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
}

function findLastRow(sheetName, startingRow, startingColumn){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetToSearch = spreadsheet.getSheetByName(sheetName);
  var count = 0;
  Logger.log("searching last row, starting from row " + startingRow);
  Logger.log(`first cell: ${startingColumn}${startingRow}`);
  
  for(i = startingRow; sheetToSearch.getRange(`${startingColumn}${i}`).getValue() != ""; i++){
    Logger.log("row value: " + sheetToSearch.getRange(`${startingColumn}${i}`).getValue());
    count++;
  }
  Logger.log("sheet last row: " + Number(startingRow + count));
  return startingRow + count;
}

function avgStockPrice() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var activeCell = activeSheet.getActiveCell();
  var activeCellRow = activeCell.getRow();
  var activeCellColumn = activeCell.getColumn();

  var stockTicker = activeSheet.getRange(activeCellRow, 3).getValue();
  
  var actionsSheet = sheet.getSheetByName("פעולות");
  var lrow = actionsSheet.getLastRow();
  var frow = Number(props.getProperty("actionPagePurchaseFirstRow"));
  for(i = frow;i <= lrow ;i++){
    if(actionsSheet.getRange("A"+i).getValue() == 'סה"כ'){
      lrow = i - 1;
      props.setProperty("actionPagePurchaseLastRow", lrow)
      break;
    }
  }

  var summarySheet = sheet.getSheetByName("תמונת מצב");
  var weightedAvgRange = activeSheet.getRange(props.getProperty("weightedAvgFirstRow"))
  for(i = weightedAvgRange.getRow(); i <= summarySheet.getLastRow(); i++){
    var currentCell = summarySheet.getRange("R"+i)
    Logger.log(`row: ${i}, current cell value: ${currentCell.getValue()}`)
    if(currentCell.getValue() == "EOR"){
      break;
    }
      var stockTicker = summarySheet.getRange(i, 3).getValue();
      Logger.log(`stock ticker: ${stockTicker}`)
      var stockUnitCount = summarySheet.getRange(`M${i}`)
      stockUnitCount.setFormula(`=IF(G${i} = "", "", SUMIF('פעולות'!$D${frow}:$D${lrow}, C${i},'פעולות'!$G${frow}:$G${lrow}))`)
      if(summarySheet.getRange(`F${i}`).getValue() == 'ארה"ב'){
        currentCell.setFormula(`=IFERROR(AVERAGE.WEIGHTED(QUERY('פעולות'!$A${frow}:$Q${lrow},\"select H where B='קניה' and D='${stockTicker}'\",1),(QUERY  ('פעולות'!$A${frow}:$Q${lrow},\"select G where B='קניה' and D='${stockTicker}'\",1))))`)
        currentCell.setNumberFormat('$#,##0.00')
      }
      else if(summarySheet.getRange(`F${i}`).getValue() == 'ישראל'){
        currentCell.setFormula(`=IFERROR(AVERAGE.WEIGHTED(QUERY('פעולות'!$A${frow}:$Q${lrow},\"select H where B='קניה' and D='${stockTicker}'\",1),(QUERY  ('פעולות'!$A${frow}:$Q${lrow},\"select G where B='קניה' and D='${stockTicker}'\",1)))/100)`)
        currentCell.setNumberFormat('₪#,##0.00')
      }
  }

  if(activeCellColumn == 6){
    if(activeSheet.getRange(activeCellRow, activeCellColumn).getValue() == 'ארה"ב'){
      activeSheet.getRange(activeCellRow,18).setNumberFormat('$#,##0.00##')
    }
    else if(activeSheet.getRange(activeCellRow, activeCellColumn).getValue() == 'ישראל'){
      activeSheet.getRange(activeCellRow,18).setNumberFormat('₪#,##0.00##')
    }
  }
}

function cryptoCoinsCount(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var cryptoSheet = sheet.getSheetByName(props.getProperty("cryptoActionsPage"));
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var activeCell = activeSheet.getActiveCell();
  var activeCellRow = activeCell.getRow();
  var activeCellColumn = activeCell.getColumn();

  var lrow = cryptoSheet.getLastRow();
  var frow = Number(props.getProperty("cryptoActionsFirstRow"));
  for(var i = frow;i <= lrow ;i++){
    if(cryptoSheet.getRange("A"+i).getValue() == 'סה"כ'){
      lrow = i - 1;
      props.setProperty("cryptoActionsLastRow", lrow)
      break;
    }
  }

  var cryptoSummaryFirstRow = Number(props.getProperty("cryptoSummaryFirstRow"))
  for(var i = cryptoSummaryFirstRow; i <= activeSheet.getLastRow(); i++){
    var cryptoSymbol = activeSheet.getRange(i, 11);
    var cellToSetFormula = activeSheet.getRange(i,15);
    if(cryptoSymbol.isBlank()){
      return
    }
    Logger.log(`=SUMIF('קריפטו פעולות'!$D$${frow}:$D$${lrow},$K$${i},'קריפטו פעולות'!$E$${frow}:$E$${lrow})+SUMIF('קריפטו פעולות'!$H$${frow}:$H$${lrow},$K$${i},'קריפטו פעולות'!$I$${frow}:$I$${lrow})`)
    cellToSetFormula.setFormula(`=SUMIF('קריפטו פעולות'!$D$${frow}:$D$${lrow},$K$${i},'קריפטו פעולות'!$E$${frow}:$E$${lrow})+SUMIF('קריפטו פעולות'!$H$${frow}:$H$${lrow},$K$${i},'קריפטו פעולות'!$I$${frow}:$I$${lrow})`)
  }

}
// =SUMIF('קריפטו פעולות'!D$20:D$48, K19, 'קריפטו פעולות'!E$20:E$48)+SUMIF('קריפטו פעולות'!H$20:H$48, K19,'קריפטו פעולות'!I$20:I$48)

function currencyFormat(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (e.range.getSheet().getName() === 'פעולות'){
    var activeCell = SpreadsheetApp.getActive().getActiveCell();
    var activeCellRow = activeCell.getRow();
    if(activeCell.getColumn() == 10 || activeCell.getColumn() == 9){
      if(sheet.getRange(`I${activeCellRow}`).getValue() == 'USD'){
        sheet.getRange(activeCellRow,10).setNumberFormat('$#,##0.00##')
      }
      else if(sheet.getRange(`I${activeCellRow}`).getValue() == 'ILS'){
        sheet.getRange(activeCellRow,10).setNumberFormat('₪#,##0.00##')
      }
    } 
  }
}


// function onOpen() 
// {
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('GetValues')
//       .addItem('currency format', 'currencyFormat')
//       .addToUi();
// }

function onEdit(e) {      
  var activeCell = SpreadsheetApp.getActive().getActiveCell();

  if (e.range.getSheet().getName() === 'תמונת מצב') {
    if (e.range.getA1Notation() === 'F'+activeCell.getRow()) {
      SpreadsheetApp.getActive().toast("running avgStockPrice on cell "+activeCell.getA1Notation(),"running script")
      avgStockPrice()
    }
  }
  if (e.range.getSheet().getName() === 'פעולות') {
    avgStockPrice()
  }

  currencyFormat(e)
}

// --- Freeze formula cells to values (avoids #N/A breaking downstream) ---
var FREEZE_ERROR_DISPLAY_VALUES = ["#N/A", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#ERROR!"];

function isFreezeErrorDisplay(displayVal) {
  if (displayVal == null || displayVal === "") return true;
  var s = (displayVal + "").trim();
  for (var i = 0; i < FREEZE_ERROR_DISPLAY_VALUES.length; i++) {
    if (s === FREEZE_ERROR_DISPLAY_VALUES[i]) return true;
  }
  return false;
}

/**
 * For each cell in range: if it has a formula and the displayed value is not an error and not blank,
 * replace the formula with that value (freeze to value).
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @returns {number} count of cells frozen
 */
function freezeRangeToValues(range) {
  if (!range || range.getNumRows() === 0 || range.getNumColumns() === 0) return 0;
  var formulas = range.getFormulas();
  var displayValues = range.getDisplayValues();
  var values = range.getValues();
  var frozen = 0;
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      var formula = (formulas[i][j] || "").toString().trim();
      var disp = displayValues[i][j];
      if (formula !== "" && !isFreezeErrorDisplay(disp) && (disp + "").trim() !== "") {
        range.getCell(i + 1, j + 1).setValue(values[i][j]);
        frozen++;
      }
    }
  }
  return frozen;
}

/**
 * Freeze the currently selected cells to values (formula -> value when result is valid).
 * Run from menu: Special functions -> Freeze selected cells to values.
 */
function freezeSelectedCellsToValues() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getActive().toast("Select one or more cells first.", "Freeze to values", 3);
    return;
  }
  var frozen = freezeRangeToValues(range);
  SpreadsheetApp.getActive().toast("Froze " + frozen + " cell(s) to values.", "Freeze to values", 3);
}

/**
 * Freeze a block (several columns, from first row to last row) to values.
 * Defaults: sheet "תמונת מצב", columns J:L, first row 52. Edit this function to change.
 * Run from menu: Special functions -> Freeze portfolio block to values.
 */
function freezeBlockToValues() {
  var sheetName = props.getProperty("summaryPage") || "תמונת מצב";
  var columnLetters = "J:L";
  var firstRow = 52;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getActive().toast('Sheet "' + sheetName + '" not found.', "Freeze block", 4);
    return;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < firstRow) {
    SpreadsheetApp.getActive().toast("No data rows in block.", "Freeze block", 3);
    return;
  }
  var parts = columnLetters.split(":");
  var startCol = parts[0];
  var endCol = parts[1] || parts[0];
  var rangeA1 = startCol + firstRow + ":" + endCol + lastRow;
  var range = sheet.getRange(rangeA1);
  var frozen = freezeRangeToValues(range);
  SpreadsheetApp.getActive().toast("Froze " + frozen + " cell(s) in block.", "Freeze block", 3);
}