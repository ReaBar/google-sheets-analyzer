var scriptProperties = PropertiesService.getScriptProperties();
var momentLoaded = false;
var props = PropertiesService.getUserProperties();
props.setProperty('rcringTransFirstRow', 9);
props.setProperty('transFirstRow', 9);
//props.setProperty('currentTransactionsSheet', 'Current Transactions');
props.setProperty('allTransactionsSheet', 'Transactions');
props.setProperty('recurringTransactionsSheet', 'Recurring Transactions');
props.setProperty('categoryTransfersSheet', 'Category Transfers');
props.setProperty('recurringCategoryTransfersSheet','Recurring Category Transfers')
props.setProperty('categoryReportsSheet', 'Category Reports');
props.setProperty('netWorthSheet', 'Net Worth Reports')
props.setProperty('oneZeroDebitCardAccount', '💳 DC One Zero (Rea) 6170')
props.setProperty('oneZeroCheckingAccount', '💰 Checking (One Zero)')

function loadMomentJS() {
    if (!momentLoaded) {
        eval(UrlFetchApp.fetch("https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js").getContentText());
        momentLoaded = true;
    }
}


// function setPreviousTransactionLastRow(rowNum){
//   scriptProperties.setProperty('prvsTransLastRow', rowNum);
// }

// function getPreviousTransactionLastRow(){
//   return Number(scriptProperties.getProperty('prvsTransLastRow'));
// }

function setTransactionLastRow(rowNum){
  if(rowNum < props.getProperty("transFirstRow")){
    scriptProperties.setProperty('transLastRow', props.getProperty("transFirstRow"))
  } else {
    scriptProperties.setProperty('transLastRow', rowNum);
  }
}

function getTransactionLastRow(){
  return Number(scriptProperties.getProperty('transLastRow'));
}

function setCategoryTransfersLastRow(rowNum){
  scriptProperties.setProperty('categoryTransfersLastRow', rowNum);
}

function getCategoryTransfersLastRow(){
  return Number(scriptProperties.getProperty('categoryTransfersLastRow'));
}

function setNetWorthLastRow(rowNum){
  scriptProperties.setProperty('netWorthLastRow', rowNum);
}

function getNetWorthLastRow(){
  return Number(scriptProperties.getProperty('netWorthLastRow'));
}


function test_EventBasedFunction(){
  var debug_e = {
    'day-of-month': 1,
    'month': 10,
    'year': 2022
  };
  // recurringTransactions(debug_e)
  // moveToPreviousTransactions(debug_e)
  // clearPreviousMonthTransactions(debug_e);
  // recurringCategoryTransfers(debug_e)
  copyCategoryReportsTotalSumFormula(debug_e);
}

function test_onEdit() {
  onEdit({
    user : Session.getActiveUser().getEmail(),
    source : SpreadsheetApp.getActiveSpreadsheet(),
    range : SpreadsheetApp.getActiveSpreadsheet().getRange("B129:I129"),
    value : SpreadsheetApp.getActiveSpreadsheet().getRange("B129:I129").getValues(),
    authMode : "LIMITED",
    oldValue: "5101b774-404f-4994-aa78-b910355c823e"
  });
}

function test_onEdit_json() {
  // onEdit({"authMode":"LIMITED","value":"✅","oldValue":"🅿️","range":{"columnEnd":9,"columnStart":9,"rowEnd":129,"rowStart":129}, "source": {Spreadsheet_object},"user":{"email":"rea.bar@gmail.com","nickname":"rea.bar"}})
  var e = {
    "authMode": "LIMITED",
    "range": {
      "columnStart": 8,
      "rowStart": 129,
      "rowEnd": 129,
      "columnEnd": 8
    },
    "source": {Spreadsheet_object},
    "oldValue": "🅿️",
    "user": {
      "nickname": "rea.bar",
      "email": "rea.bar@gmail.com"
    },
    "value": "✅"
  };
  onEdit(e)
}

function test_onChange() {
  onMyChange({
    user : Session.getActiveUser().getEmail(),
    source : SpreadsheetApp.getActiveSpreadsheet(),
    range : SpreadsheetApp.getActiveSpreadsheet().getRange("B83:I83"),
    value : SpreadsheetApp.getActiveSpreadsheet().getRange("B83:I83").getValues(),
    authMode : "FULL",
    oldValue: "5101b774-404f-4994-aa78-b910355c823e"
  });
}

function test_handleDebitCardTransactions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(props.getProperty("allTransactionsSheet"))
  var row = 1923
  handleDebitCardTransactions(sheet, row)
}

function test_prvsTransactionsLastRow() {
  transactionsLastRow(SpreadsheetApp.getActiveSpreadsheet());
}

function test_sortFunction(){
  triggeredSheetSort();
  // var sheet = SpreadsheetApp.getActiveSpreadsheet();
  // var range = "B31:I32";
  // sortTransactionsRange(sheet, range);
}

function test_copyRange(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('recurringTransactionsSheet'));
  var allTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('allTransactionsSheet'));
  copyRange(currentTransactionsSheet, allTransactionsSheet, 40);
}

function test_clearRowData(){
  clearRowData('Transactions', 131)
}

function onChange() {
  ScriptApp.newTrigger("onMyChange")
   .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
   .onChange()
   .create();
}

function onMyChange(e) {
  try {
    Logger.log("Handling change event of type: " + e.changeType);
    // if (e.changeType && e.changeType === "EDIT") {
    // // Skip processing because edits are handled by the onEdit trigger
    // return;
    // }
    var transFirstRow = Number(props.getProperty("transFirstRow"));
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = e.source.getActiveSheet();
    Logger.log("Processing on sheet: " + sheet.getName());

    // if (sheet.getName() == props.getProperty("currentTransactionsSheet")) {
    //   var range = sheet.getActiveRange();
    //   var row = range.getLastRow();
    //   var column = range.getColumn();
    //   var idCell = sheet.getRange(`I${row}`);
    //   var outFlow = sheet.getRange("C" + row);
    //   var inFlow = sheet.getRange("D" + row);
    //   var category = sheet.getRange("E" + row);
    //   var account = sheet.getRange("F" + row);
    //   var isDataRangeNotBlank = false;

    //   if ((!outFlow.isBlank() || !inFlow.isBlank()) && !category.isBlank() && !account.isBlank()) {
    //     isDataRangeNotBlank = true;
    //   }

    //   if (isDataRangeNotBlank && idCell.isBlank() && transFirstRow <= row) {
    //     var allTransactionsSheet = spreadsheet.getSheetByName(props.getProperty("allTransactionsSheet"));
    //     var uniqueId = uuid();
    //     idCell.setValue(uniqueId);
    //     var sourceRange = sheet.getRange(`B${row}:I${row}`);
    //     var sourceValues = sourceRange.getValues();
    //     var prvsTransRowToEdit = previousTransactionsLastRow(allTransactionsSheet);
    //     var targetRange = allTransactionsSheet.getRange(`B${prvsTransRowToEdit}:I${prvsTransRowToEdit}`);
    //     targetRange.setValues(sourceValues);
    //   }
    // }

    if (sheet.getName() == props.getProperty("recurringTransactionsSheet")) {
      var range = sheet.getActiveRange();
      var row = range.getRow();
      var column = range.getColumn();
      var cell = range.getA1Notation();
      Logger.log(`${range.getValue()}, column: ${column}`);

      if (`${column}` == 7 && range.getValue() > 0) {
        var alternateRange = sheet.getRange(`B${row}:G${row}`);
        if (row % 2 == 0) {
          alternateRange.setBackground("#F3F3F3");
        } else {
          alternateRange.setBackground("#FFFFFF");
        }
      }
    }

    // ✅ Trigger updateShowHideColumn ONLY for allTransactionsSheet
    if (sheet.getName() == props.getProperty("allTransactionsSheet")) {
      
      loadMomentJS(); // Load Moment.js only once per script execution
      Logger.log("MomentJS loaded.");

      var range = sheet.getActiveRange();
      var row = range.getRow();
      var column = range.getColumn();
      Logger.log("Active range: row " + row + ", column " + column);

      // Retrieve and log the values for columns I and B on this row.
      var colICell = sheet.getRange("I" + row);
      var colBCell = sheet.getRange("B" + row);
      Logger.log("Cell I" + row + " is blank? " + colICell.isBlank());
      Logger.log("Cell B" + row + " is blank? " + colBCell.isBlank());
      Logger.log("Row " + row + " < transFirstRow (" + props.getProperty('transFirstRow') + ")? " + (row < Number(props.getProperty('transFirstRow'))));

      // If cell I isn't blank, cell B is blank, or the row is before transFirstRow, exit.
      if (!colICell.isBlank() || colBCell.isBlank() || row < Number(props.getProperty('transFirstRow'))) {
        Logger.log("Exiting because one of the conditions was met (I not blank, or B blank, or row too low).");
        return;
      }

      // Generate and set a unique ID in column I.
      var uniqueId = uuid();
      Logger.log("Generated uniqueId: " + uniqueId);
      colICell.setValue(uniqueId);

      // Process the description in column G.
      var description = sheet.getRange("G" + row).getValue().toString();
      Logger.log("Original description in G" + row + ": " + description);
      var searchString = " - Added from Aspire Android app";
      if (description.indexOf(searchString) !== -1) {
        Logger.log("Search string found in description. Removing it.");
        var newDescription = description.replace(searchString, "");
        Logger.log("New description: " + newDescription);
        sheet.getRange("G" + row).setValue(newDescription);
      } else {
        Logger.log("Search string not found in description. No change made.");
      }

      // Use Moment.js to get dates.
      var currentDate = moment().startOf("month");
      Logger.log("Current date (start of month): " + currentDate.format("DD/MM/YYYY"));
      
      var cellDateValue = sheet.getRange("B" + row).getValue();
      Logger.log("Cell B" + row + " date value: " + cellDateValue);
      
      var cellDate = moment(cellDateValue, "DD/MM/YYYY");
      Logger.log("Parsed cell date: " + cellDate.format("DD/MM/YYYY"));
      
      var cellDateStartOfMonth = cellDate.startOf("month");
      Logger.log("Cell date start of month: " + cellDateStartOfMonth.format("DD/MM/YYYY"));
      
      // Call updateShowHideColumn and log the action.
      Logger.log("Calling updateShowHideColumn for row " + row);
      updateShowHideColumn(sheet, row);
      
      // Call handleDebitCardTransactions and log the action.
      Logger.log("Calling handleDebitCardTransactions for row " + row);
      handleDebitCardTransactions(sheet, row);
      
      Logger.log("Processing complete for row " + row);
    }
  } catch (error) {
    showAndThrow_(error);
  }
}


function handleDebitCardTransactions(sheet, row){
  var account = sheet.getRange("F"+row).getValue()
  var category = sheet.getRange("E"+row).getValue()
  var uniqueId = sheet.getRange("I"+row).getValue()
  var status = sheet.getRange("H"+row).getValue()
  var isOutflowTransaction = !sheet.getRange("C"+row).isBlank()
  var amount = !sheet.getRange("C"+row).isBlank() ? sheet.getRange("C"+row).getValue() : sheet.getRange("D"+row).getValue()
  if(!account.includes(props.getProperty('oneZeroDebitCardAccount')) || category.includes("Account Transfer")){
    return
  }
  var uniqueIdsRanges = sheet.createTextFinder(uniqueId).findAll();
  var ranges = uniqueIdsRanges.map(r => ({row: r.getRow(), column: r.getColumn()}))
  if (ranges.length > 1) {
    Logger.log(`Updating debit card transactions to amount ${amount}, found related rows: ${JSON.stringify(ranges)}`)
    ranges.forEach(function (value) {
      var isOutflowTransaction = !sheet.getRange("C"+value.row).isBlank()
      if (value.column == 7.0 && isOutflowTransaction) {
        sheet.getRange("C" + value.row).setValue(amount)
      } else if(value.column == 7) {
        sheet.getRange("D" + value.row).setValue(amount)
      }
      sheet.getRange("H" + value.row).setValue(status)
    })
  } else {
    var oneZeroDebitCardAccount = props.getProperty('oneZeroDebitCardAccount')
    var oneZeroCheckingAccount = props.getProperty('oneZeroCheckingAccount')
    if (isOutflowTransaction) {
      createAccountTransferTransactions(oneZeroCheckingAccount, oneZeroDebitCardAccount, sheet, row, amount)
    } else {
      createAccountTransferTransactions(oneZeroDebitCardAccount, oneZeroCheckingAccount, sheet, row, amount)
    }
  }
}

function copyCategoryReportsTotalSumFormula(e) {
  var month = e['month']
  var year = e['year']
  var categoryReportsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(props.getProperty("categoryReportsSheet"))
  var columnToEdit = 7 + month
  var yearViewCell = "B4"
  categoryReportsSheet.getRange(yearViewCell).setValue(year)
  if(month==1){
    console.log("Currently can't handle first month as formulas needed to be copied from previous year by row")
    return;
  }
  for(var i = 25; !categoryReportsSheet.getRange(`B${i}`).isBlank(); i++ ){
    if(!isCellCategory(props.getProperty("categoryReportsSheet"), `B${i}`)){
      continue;
    }
    var formula = categoryReportsSheet.getRange(i, columnToEdit-1).getFormulaR1C1();
    categoryReportsSheet.getRange(i, columnToEdit).setFormulaR1C1(formula)
  }
}

function isCellCategory(sheetName, cellA1Notation) {
  var range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(cellA1Notation);
  Logger.log(`range background: ${range.getBackground()}`)
  if(range.getBackground() == '#e5e3e1'){
    return true;
  }
  return false;
}

function createAccountTransferTransactions(fromAccount, toAccount, sheet, row, amount){
  var transactionDate = sheet.getRange("B"+row).getValue()
  var originalUUID = sheet.getRange("I"+row).getValue()
  var status = sheet.getRange("H"+row).getValue()
  var lastRow =transactionsLastRow(sheet)
  var firstTransactionUniqueId = uuid()
  var firstTransactionValues = [[transactionDate, amount, null, "↕️ Account Transfer", fromAccount, "automated 💳DC account transfer - original trx uuid " + originalUUID, status, firstTransactionUniqueId]]
  sheet.getRange(lastRow, 2, 1, 8).setValues(firstTransactionValues)
  updateShowHideColumn(sheet, lastRow)
  lastRow += 1
  var secondTransactionUniqueId = uuid()
  var secondTransactionValues = [[transactionDate, null, amount, "↕️ Account Transfer", toAccount, "automated 💳DC account transfer - original trx uuid " + originalUUID, status,    secondTransactionUniqueId]]
  sheet.getRange(lastRow, 2, 1, 8).setValues(secondTransactionValues)
  updateShowHideColumn(sheet, lastRow)
}

function copyRange(fromSheet, toSheet, fromRow){
  var sourceRange = fromSheet.getRange(`B${fromRow}:I${fromRow}`);
  var sourceValues = sourceRange.getValues();
  var prvsTransRowToEdit = transactionsLastRow(toSheet);
  var targetRange = toSheet.getRange(`B${prvsTransRowToEdit}:I${prvsTransRowToEdit}`);
  targetRange.setValues(sourceValues);
}

// function copyRangeToCurrentTransactions(fromSheet, toSheet, fromRow){
//   var sourceRange = fromSheet.getRange(`B${fromRow}:I${fromRow}`);
//   var sourceValues = sourceRange.getValues();
//   var currentTransRowToEdit = transactionsLastRow(toSheet);
//   var targetRange = toSheet.getRange(`B${currentTransRowToEdit}:I${currentTransRowToEdit}`);
//   targetRange.setValues(sourceValues);
// }

// Triggered daily - between 02:00-03:00
function recurringTransactions(e) {
  var currentDateOfTheMonth = e['day-of-month'];
  Logger.log("current date: " + currentDateOfTheMonth);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var recurringTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('recurringTransactionsSheet'));
 // var currentTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('currentTransactionsSheet'));
  var allTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('allTransactionsSheet'));
  var rcringFirstRow = Number(props.getProperty('rcringTransFirstRow'));
  var transLastRow = transactionsLastRow(allTransactionsSheet)
  var currentDate = Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy");
  for(var i = rcringFirstRow; recurringTransactionsSheet.getRange("B"+i).getValue() <= currentDateOfTheMonth && recurringTransactionsSheet.getRange("B"+i).getValue() != ""; i++ ){
    Logger.log("recurring transactions day of month: " + recurringTransactionsSheet.getRange("B"+i).getValue());
    var recurringTransDate = recurringTransactionsSheet.getRange("B"+i).getValue();
    var installmentRecurringTrx = recurringTransactionsSheet.getRange("G"+i);
    if(recurringTransDate == currentDateOfTheMonth && (installmentRecurringTrx.isBlank() || Number(installmentRecurringTrx.getValue()) > 0)){      
      var amount = recurringTransactionsSheet.getRange("C"+i).getValue();
      var category = recurringTransactionsSheet.getRange("D"+i).getValue();
      var account = recurringTransactionsSheet.getRange("E"+i).getValue();
      var description;
      if(Number(installmentRecurringTrx.getValue()) > 0){
        description = `${recurringTransactionsSheet.getRange("F"+i).getValue()} - נותר לשלם ${Number(installmentRecurringTrx.getValue())} תשלומים כולל תשלום זה`;
      } else{
        description = recurringTransactionsSheet.getRange("F"+i).getValue();
      }
      Logger.log("setting row: " + transLastRow);
      setRecurringTransactionProperties(allTransactionsSheet, transLastRow, currentDate, amount, category, account, description);
      //copyRange(currentTransactionsSheet, allTransactionsSheet, transLastRow);
      transLastRow++;
      setTransactionLastRow(transLastRow);
      if(!installmentRecurringTrx.isBlank()){
        var remainingTrxs = Number(installmentRecurringTrx.getValue() - 1);
        if(remainingTrxs == 0){
          recurringTransactionsSheet.getRange(i, 2,1,6).setBackgroundRGB(234, 153, 153);
          recurringTransactionsSheet.getRange("G"+i).setValue(remainingTrxs);
        }
        else{
          recurringTransactionsSheet.getRange("G"+i).setValue(remainingTrxs);
        }
      }
    } else if(Number(installmentRecurringTrx.getValue()) == -1){
      recurringTransactionsSheet.deleteRow(i);
    }
  }
}

function setRecurringTransactionProperties(sheet,row,date, amount, category, account, description){
  var uniqueId = uuid();
  sheet.getRange("B"+row).setValue(date)
  if(amount < 0){
    sheet.getRange("C"+row).setValue(amount*-1)
  }
  else{
    sheet.getRange("D"+row).setValue(amount)
  }
  sheet.getRange("E"+row).setValue(category)
  sheet.getRange("F"+row).setValue(account)
  sheet.getRange("G"+row).setValue(description)
  sheet.getRange("H"+row).setValue("🅿️")
  sheet.getRange("I"+row).setValue(uniqueId)
}

// Triggered daily - between 03:00-04:00
function recurringCategoryTransfers(e) {
  var currentDateOfTheMonth = e['day-of-month'];
  Logger.log("current date: " + currentDateOfTheMonth);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var recurringCategoryTransfersSheet = spreadsheet.getSheetByName(props.getProperty('recurringCategoryTransfersSheet'));
  var categoryTransfersSheet = spreadsheet.getSheetByName(props.getProperty('categoryTransfersSheet'));
  var rcringFirstRow = Number(props.getProperty('rcringTransFirstRow'));
  var categoryTransferLastRow = categoryTransfersLastRow(categoryTransfersSheet)
  var currentDate = Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy");
  for(var i = rcringFirstRow; recurringCategoryTransfersSheet.getRange("B"+i).getValue() <= currentDateOfTheMonth && recurringCategoryTransfersSheet.getRange("B"+i).getValue() != ""; i++ ){
    Logger.log("recurring transactions day of month: " + recurringCategoryTransfersSheet.getRange("B"+i).getValue());
    var recurringTransDate = recurringCategoryTransfersSheet.getRange("B"+i).getValue();
    if(recurringTransDate == currentDateOfTheMonth){      
      var amount = recurringCategoryTransfersSheet.getRange("C"+i).getValue();
      var fromCategory = recurringCategoryTransfersSheet.getRange("D"+i).getValue();
      var toCategory = recurringCategoryTransfersSheet.getRange("E"+i).getValue();
      var description = recurringCategoryTransfersSheet.getRange("F"+i).getValue();
      Logger.log("setting row: " + categoryTransferLastRow);
      setRecurringCategoryTransfersProperties(categoryTransfersSheet, categoryTransferLastRow, currentDate, amount, fromCategory, toCategory, description);
      categoryTransferLastRow++;
      setCategoryTransfersLastRow(categoryTransferLastRow);
    }
  }
}

function setRecurringCategoryTransfersProperties(sheet,row,date, amount, fromCategory, toCategory, description){
  sheet.getRange("B"+row).setValue(date)
  if(amount < 0){
    amount = amount*-1
  }
  sheet.getRange("C"+row).setValue(amount)
  sheet.getRange("D"+row).setValue(fromCategory)
  sheet.getRange("E"+row).setValue(toCategory)
  sheet.getRange("F"+row).setValue(description)
}

// // Triggered monthly - 4th of every month between 04:00-05:00
// function clearPreviousMonthTransactions(e){
//   var currentMonth = new Date().getMonth();
//   var currentYear = new Date().getFullYear();
//   var currentDate = new Date(currentYear, currentMonth, 1);
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   var currentTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('currentTransactionsSheet'));
//   var allTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('allTransactionsSheet'));
//   var transFirstRow = Number(props.getProperty('transFirstRow'));
//   var rowsDeletedCount = 0
//   for(var i = transFirstRow; i < currentTransactionsSheet.getRange("B"+i).getValue() != "";){
//     var uniqueId = currentTransactionsSheet.getRange("I"+i).getValue()
//     var textFinder = allTransactionsSheet.createTextFinder(uniqueId);
//     if(textFinder.findAll()[0] == null){
//       var prvsTransRowToEdit = previousTransactionsLastRow(allTransactionsSheet);
//       var sourceValues = currentTransactionsSheet.getRange(i, 1, 1, 8).getValues()
//       allTransactionsSheet.getRange(prvsTransRowToEdit, 1, 1, 8).setValues(sourceValues)
//     }
//     var transactionDate = new Date(currentTransactionsSheet.getRange("B"+i).getValue());
//     if(transactionDate < currentDate){
//       currentTransactionsSheet.deleteRow(i)
//       rowsDeletedCount++;
//     }
//     else{
//       i++
//     }
//   }
//   var currentTransactionLastRow = getTransactionLastRow()
//   setTransactionLastRow(currentTransactionLastRow - rowsDeletedCount - 5)
//   Logger.log(`current transactions last row is now: ${getTransactionLastRow()}`)
// }

function clearRowData(sheetName, row){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  sheet.deleteRow(row)
}

// function verifyAllTransactionsExist(){
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   var allTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('allTransactionsSheet'));
//   var currentTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('currentTransactionsSheet'));
//   var firstRow = Number(props.getProperty('transFirstRow'));
//   for(var i = firstRow; currentTransactionsSheet.getRange("B"+i).getValue() != ""; i++){
//     var uniqueId = currentTransactionsSheet.getRange("I"+i).getValue()
//     var textFinder = allTransactionsSheet.createTextFinder(uniqueId);
//     if(textFinder.findAll()[0] == null){
//       var prvsTransRowToEdit = previousTransactionsLastRow(allTransactionsSheet);
//       var sourceValues = currentTransactionsSheet.getRange(i, 2, 1, 8).getValues()
//       allTransactionsSheet.getRange(prvsTransRowToEdit, 2, 1, 8).setValues(sourceValues)
//     }
//   }
// }

function onEdit(e) {
  var startTime = new Date().getTime(); // Record the start time
  Logger.log("onEdit started at: " + startTime);

  var sheet = e.source.getActiveSheet();
  Logger.log("Time after getting sheet: " + (new Date().getTime() - startTime) + "ms");

  if (sheet.getName() === "Transactions") {
    Logger.log(`onEdit column: ${e.range.getColumn()}`);
    
    // Case 1: if cell E4 was edited, refresh the filter.
    if (e.range.getA1Notation() === "E4") {
      refreshFilter();
      Logger.log("After refreshFilter: " + (new Date().getTime() - startTime) + "ms");
    }
    // Case 2: if column 8 was edited, update the Show/Hide column.
    else if (e.range.getColumn() === 8) {
      updateShowHideColumn(sheet, e.range.getRow());
      Logger.log("After updateShowHideColumn: " + (new Date().getTime() - startTime) + "ms");
    }
    // Case 3: if column B (2) was edited, check if column I already has a value and update Show/Hide.
    else if (e.range.getColumn() === 2) {
      var row = e.range.getRow();
      var uuidCell = sheet.getRange(row, 9); // Column I
      if (!uuidCell.getValue()) { //
        Logger.log("UUID doesn't exists at row " + row + ". creating a new one.");
        var generatedId = uuid();
        uuidCell.setValue(generatedId);
        Logger.log("After setting UUID: " + (new Date().getTime() - startTime) + "ms");
      }
      else{
        Logger.log("UUID already exists at row " + row + ".");
      }
      updateShowHideColumn(sheet, row);
    }
  }
  Logger.log("Total onEdit duration: " + (new Date().getTime() - startTime) + "ms");
}

function logTimeCheckpoint(label, startTime) {
  var elapsed = new Date().getTime() - startTime;
  Logger.log(label + ": " + elapsed + "ms");
}

// function onEdit(e) {
//   var sheet = e.source.getActiveSheet();
//   if (sheet.getName() === "Transactions") {
//     Logger.log(`onEdit column: ${e.range.getColumn()}`)
//     if(e.range.getA1Notation() === "E4"){
//       refreshFilter();
//     }
//     else if(e.range.getColumn() === 8){
//       updateShowHideColumn(sheet, e.range.getRow());
//     }
//     else if (e.range.getColumn() === 2) {
//       var row = e.range.getRow();
//       // Get the cell in column I (column index 9) for the same row
//       var uuidCell = sheet.getRange(row, 9);
//       if (!uuidCell.getValue()) {
//         var generatedId = uuid();
//         uuidCell.setValue(generatedId);
//       }
//     }
//   }
  // if (e.source.getActiveSheet().getName() == props.getProperty('currentTransactionsSheet')){
  //   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //   var allTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('allTransactionsSheet'));
  //   var currentTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('currentTransactionsSheet'));
  //   var range = e.range;
  //   var cell = currentTransactionsSheet.getRange(e.range.getA1Notation());
  //   var column;
  //   var columnStart = range.columnStart;
  //   var columnEnd = range.columnEnd;
  //   var row = range.getRow();
  //   var idCell = currentTransactionsSheet.getRange("I"+row);

  //   var id;
  //   var prvsTransRowToEdit;
  //   if(idCell.isBlank()){
  //     id = uuid();
  //     Logger.log(`generated uuid is: ${id}`)
  //   }
  //   else{
  //     id = idCell.getValue();
  //     var textFinder = allTransactionsSheet.createTextFinder(id);
  //     if(textFinder.findAll()[0] == null){
  //       prvsTransRowToEdit = previousTransactionsLastRow(allTransactionsSheet);
  //     }
  //     else{
  //      prvsTransRowToEdit  = textFinder.findAll()[0].getRow();
  //     }
  //   }

  //   if(columnStart == columnEnd){
  //     column = columnStart;
  //   }

  //   Logger.log(`editing row ${row}, column ${column}`)

  //   if(columnStart != columnEnd || column == 8){
  //     SpreadsheetApp.getActive().toast("editing more than one row or updated status");
  //     var sourceRange = currentTransactionsSheet.getRange(`B${row}:I${row}`);
  //     var sourceValues = sourceRange.getValues();
  //     var targetRange = allTransactionsSheet.getRange(`B${prvsTransRowToEdit}:I${prvsTransRowToEdit}`);
  //     targetRange.setValues(sourceValues);
  //     handleDebitCardTransactions(allTransactionsSheet, prvsTransRowToEdit)
  //   }

  //   else if(column == 2){
  //     if(idCell.isBlank() && !currentTransactionsSheet.getRange(`B${row}`).isBlank()){
  //       idCell.setValue(id);
  //     }
  //     var previousTransIdCell = allTransactionsSheet.getRange("I"+prvsTransRowToEdit);
  //     if(previousTransIdCell.isBlank()){
  //       previousTransIdCell.setValue(idCell.getValue());
  //     }
  //     SpreadsheetApp.getActive().toast("copying cell: "+column +" to all transactions sheet");
  //     var prvsTransCell = allTransactionsSheet.getRange(prvsTransRowToEdit, column);
  //     prvsTransCell.setValue(e.value);
  //     prvsTransCell.setNumberFormat("dd/mm/yyyy");
  //   }
  //   else if(3 <= column && column <= 7){
  //     Logger.log(`status: ${!currentTransactionsSheet.getRange(`H${row}`).isBlank()}`)
  //     if(!currentTransactionsSheet.getRange(`H${row}`).isBlank()){
  //       SpreadsheetApp.getActive().toast("copying cell: "+column +" to all transactions sheet");
  //       if(allTransactionsSheet.getRange("F"+prvsTransRowToEdit).getValue().includes(props.getProperty('oneZeroDebitCardAccount'))){
  //         showMessage_(`Changing account from DEBIT CARD requires removing also the account transfer transaction from ${props.getProperty('allTransactionsSheet')} sheet`, 10)
  //       }
  //       var prvsTransCell = allTransactionsSheet.getRange(prvsTransRowToEdit, column);
  //       prvsTransCell.setValue(e.value);
  //       handleDebitCardTransactions(allTransactionsSheet, prvsTransRowToEdit)
  //     }
  //   }
  //   else if(column == 8){
  //     var sourceRange = currentTransactionsSheet.getRange(`B${row}:I${row}`);
  //     var sourceValues = sourceRange.getValues();
  //     var targetRange = allTransactionsSheet.getRange(`B${prvsTransRowToEdit}:I${prvsTransRowToEdit}`);
  //     targetRange.setValues(sourceValues);
  //     handleDebitCardTransactions(allTransactionsSheet, prvsTransRowToEdit)
  //   }
  // }
//}

function updateShowHideColumn(sheet, row) {
    // Ensure we are modifying allTransactionsSheet
    if (sheet.getName() !== props.getProperty("allTransactionsSheet")) return;

    var e4 = sheet.getRange("E4").getValue(); // Get filter value from E4
    var bCell = sheet.getRange(row, 2).getValue(); // Column B (Date column)
    var jCell = sheet.getRange(row, 10); // Column J (Show/Hide column)

    // Ensure there is a valid date in column B before applying formula
    if (!bCell && e4 !== "") {  
        jCell.setValue("Hide"); // If B is empty but E4 has a value, set to Hide
    } else {  
        var formula = `=IF(AND(B${row}="", NOT(ISBLANK($E$4))), "Hide", 
                            IF(OR(B${row}="", ISBLANK($E$4), 
                                  (YEAR(B${row}) = YEAR($E$4)) * 
                                  (MONTH(B${row}) = MONTH($E$4))), 
                               "Show", "Hide"))`;
        jCell.setFormula(formula);
    }
}

function transactionsLastRow(sheet){
  var lastRow = Number(getTransactionLastRow()) - 15;
  var firstRow = props.getProperty('transFirstRow');
  Logger.log(`First row: ${firstRow}, last row: ${lastRow}`)
  if(lastRow === undefined || lastRow < firstRow){
    lastRow = Number(props.getProperty('transFirstRow'));
  }
  var count = 0;
  for(var i = lastRow; sheet.getRange("B"+i).getValue() != ""; i++){
    count++;
  }
  var newLastRow = lastRow + count;
  setTransactionLastRow(newLastRow);
  return newLastRow;
}

function categoryTransfersLastRow(categoryTransfersSheet){
  var lastRow = Number(getCategoryTransfersLastRow()) - 6;
  var firstRow = props.getProperty('transFirstRow');
  if(lastRow === undefined || lastRow < firstRow){
    lastRow = Number(props.getProperty('transFirstRow'));
  }
  var count = 0;
  for(var i = lastRow; categoryTransfersSheet.getRange("B"+i).getValue() != ""; i++){
    count++;
  }
  var newLastRow = lastRow + count;
  setCategoryTransfersLastRow(newLastRow);
  return newLastRow;
}

function triggeredSheetSort(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  //var currentTransactionsSheet = sheet.getSheetByName(props.getProperty('currentTransactionsSheet'));
  var allTransactionsSheet = sheet.getSheetByName(props.getProperty('allTransactionsSheet'));

  //var transLastRow = transactionsLastRow(currentTransactionsSheet);
  transactionsLastRow(allTransactionsSheet); //the last row is saved within the function so there is no use to the returning variable
  //var transactionsRange = `B${Number(props.getProperty('transFirstRow'))}:I${transLastRow}`;
  var previousTransactionsRange = `B${Number(props.getProperty('transFirstRow'))}:I${getTransactionLastRow()}`;

  //sortTransactionsRange(currentTransactionsSheet, transactionsRange);
  sortTransactionsRange(allTransactionsSheet, previousTransactionsRange);
}

function sortTransactionsRange(sheet, range){
  var dateColumn = 2;
  var outFlowColumn = 3;
  var inFlowColumn = 4;
  var statusColumn = 8; 
  var tableRange = range;

  var range = sheet.getRange(tableRange);
  range.sort( [{ column : dateColumn }, {column: outFlowColumn, ascending: false}, {column: inFlowColumn,ascending: false}, {column: statusColumn}] );
}

function handleFXTransactions(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var allTransactionsSheet = sheet.getSheetByName(props.getProperty('allTransactionsSheet'));
  var lastRow = Number(getTransactionLastRow()) - 50;
  for(var i = lastRow; !allTransactionsSheet.getRange("B"+i).isBlank(); i++){
    var description = allTransactionsSheet.getRange("G"+i).getValue().toString();
    var amountCell = allTransactionsSheet.getRange("C"+i)
    if(amountCell.isBlank() && description.includes("FX")){
      var regex = RegExp("\\d+(\\.\\d{1,2})?$", "gim")
      var a = regex.test(description)
      var amount = description.match("\\d+(\\.\\d{1,2})?$")[0]
      if(description.includes("$")){
        amountCell.setFormula(`=INDEX(GOOGLEFINANCE("CURRENCY:USDILS", "price", B${i}), 2, 2)*${amount}`)
      } 
      else if(description.includes("€")){
        amountCell.setFormula(`=INDEX(GOOGLEFINANCE("CURRENCY:EURILS", "price", B${i}), 2, 2)*${amount}`)
      }
      //overrideTransaction(i)
    }
  }
}

// function overrideTransaction(sourceRow){
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   var allTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('allTransactionsSheet'));
//   var currentTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('currentTransactionsSheet'));
//   var uniqueId = currentTransactionsSheet.getRange("I"+sourceRow).getValue()
//   var textFinder = allTransactionsSheet.createTextFinder(uniqueId);
//   if(textFinder.findAll()[0] != null){
//     var allTransactionsCellToOverride = textFinder.findAll()[0].getRow();
//     var sourceValues = currentTransactionsSheet.getRange(sourceRow, 2, 1, 8).getValues()
//     allTransactionsSheet.getRange(allTransactionsCellToOverride, 2, 1, 8).setValues(sourceValues)
//   }
// }

// function removeHardCodedMobileDescription(){
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   var currentTransactionsSheet = spreadsheet.getSheetByName(props.getProperty('currentTransactionsSheet'));
//   var description = currentTransactionsSheet.getRange("G44").getValue().toString();
//   var firstIndex = description.indexOf(" - Added from Aspire Android app")
//   Logger.log(`first index: ${firstIndex}`)
//   var newDescription = description.substr(0, firstIndex)
//   Logger.log(`new description: ${newDescription}`)
// }

function uuid() {
  return Utilities.getUuid();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Functions')
      .addItem('Sort sheet', 'triggeredSheetSort')
      .addItem('Convert FX transactions', 'handleFXTransactions')
      .addItem('Backfill Show/Hide column', 'backfillShowHideColumn')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))
      .addToUi();
}

function showAndThrow_(error) {
  // version 1.0, written by --Hyde, 16 April 2020
  //  - initial version
  var stackCodeLines = String(error.stack).match(/\d+:/);
  if (stackCodeLines) {
    var codeLine = stackCodeLines.join(', ').slice(0, -1);
  } else {
    codeLine = error.stack;
  }
  showMessage_(error.message + ' Code line: ' + codeLine, 30);
  throw error;
}

function getNetWorth(){
  var currentDate = Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy");
  var netWorthSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(props.getProperty('netWorthSheet'));
  var investmentsSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1S8ploU9ZuZQGZ7B1AbH8lZk9SIFe-yqPdwOSxH1GH7U/edit#gid=997506131");
  var reaHishtalmutAmount = investmentsSheet.getRangeByName('worth_hishtalmut').getValue();
  var reaPensionAmount = investmentsSheet.getRangeByName('worth_pension_estimation').getValue();
  var reaKupatGemel = investmentsSheet.getRangeByName('worth_kupat_gemel').getValue();
  var stocksWorth = Number(investmentsSheet.getRangeByName('worth_ib_stocks').getValue());
  var stocksAccountCashSheqel = Number(investmentsSheet.getRangeByName('trading_account_balance_sheqel').getValue());
  var stocksAccountCashUSDConverted = Number(investmentsSheet.getRangeByName('trading_account_balance_usd_converted').getValue());
  var cashWorth = investmentsSheet.getRangeByName('worth_cash').getValue();
  var realEstateWorth = investmentsSheet.getRangeByName('worth_real_estate').getValue();
  var cryptoWorth = investmentsSheet.getRangeByName('worth_crypto').getValue();
  var creditCardsTotal = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('total_credit_cards').getValue();
  var netWorthLastRow = getNetWorthSheetLastRow(netWorthSheet);
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,'','Otsar Hahayal Checking']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,'','One Zero Checking']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,4).setValues([[currentDate,stocksWorth + stocksAccountCashSheqel + stocksAccountCashUSDConverted,'Stocks Portfolio Account', 'meitav dash current value']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,cashWorth,'Savings']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,reaPensionAmount,'Rea\'s Pension']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,reaHishtalmutAmount,'Rea\'s Keren Hishtalmut']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,reaKupatGemel,'Rea\'s kupat gemel']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,realEstateWorth,'Home Value']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,creditCardsTotal,'Credit Card']])
  // netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,stocksWorth,'Meitav Trade Total Sum']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,cryptoWorth,'Crypto']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,'','Analyst Kupat Gemel Le\'ashkaa']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,'','Meitav Kupat Gemel Le\'ashkaa']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,3).setValues([[currentDate,'','Mortgage']])
  netWorthSheet.getRange(netWorthLastRow++, 2,1,5).setValues([[currentDate,'','','','*️⃣']])
}

function getNetWorthSheetLastRow(sheet){
  var lastRow = getNetWorthLastRow();
  if(lastRow === undefined || lastRow < 20){
    lastRow = 20
  }
  var count = 0;
  for(var i = lastRow; !sheet.getRange("B"+i).isBlank(); i++){
    count++;
  }
  lastRow = count + lastRow;
  setNetWorthLastRow(lastRow)
  return lastRow;
}

function backfillShowHideColumn() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(props.getProperty("allTransactionsSheet"));
    if (!sheet) return;
    var firstRow = Number(props.getProperty("transFirstRow"));
    var lastRow = sheet.getLastRow();
    if (lastRow < firstRow) return;
    var count = 0;
    for (var row = firstRow; row <= lastRow; row++) {
        var jCell = sheet.getRange(row, 10);
        var formula = jCell.getFormula();
        if (!formula || formula.indexOf("=") !== 0) {
            updateShowHideColumn(sheet, row);
            count++;
        }
    }
    if (count > 0) {
        showMessage_("Backfilled Show/Hide formula in " + count + " row(s).", 5);
    } else {
        showMessage_("No rows needed backfill.", 3);
    }
}

function refreshFilter() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(props.getProperty("allTransactionsSheet"));
    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 10);
    if (lastRow < 8) return;

    var filter = sheet.getFilter();
    if (filter) filter.remove();

    var filterRange = sheet.getRange(8, 1, lastRow, lastCol);
    filter = filterRange.createFilter();

    var criteria = SpreadsheetApp.newFilterCriteria()
        .setHiddenValues(["Hide"])
        .build();
    filter.setColumnFilterCriteria(10, criteria);
}

/**
* Shows a message in a pop-up.
*
* @param {String} message The message to show.
* @param {Number} timeoutSeconds Optional. The number of seconds before the message goes away. Defaults to 5.
*/
function showMessage_(message, timeoutSeconds) {
  // version 1.0, written by --Hyde, 16 April 2020
  //  - initial version
  SpreadsheetApp.getActive().toast(message, 'PAY ATTENTION', timeoutSeconds || 5);
}

function getHexValue(range) {
  return SpreadsheetApp.getActiveSheet().getRange(range).getBackground();
}