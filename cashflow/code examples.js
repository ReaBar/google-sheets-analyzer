// function moveToPreviousTransactions(){
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   var previousTransactionsSheet = spreadsheet.getSheetByName('Previous Transactions');
//   var transactionsSheet = spreadsheet.getSheetByName('Transactions');
//   var prvsTransLastRow = transactionsLastRow(previousTransactionsSheet)
//   var transFirstRow = Number(props.getProperty('transFirstRow'));
//   for(i = transFirstRow; i < transactionsSheet.getRange("B"+i).getValue() != ""; i++){
//     var id = transactionsSheet.getRange("A"+i).getValue(); 
//     var date = transactionsSheet.getRange("B"+i).getValue(); 
//     var outFlowAmount = transactionsSheet.getRange("C"+i).getValue();
//     var inFlowAmount = transactionsSheet.getRange("D"+i).getValue();
//     var category = transactionsSheet.getRange("E"+i).getValue();
//     var account = transactionsSheet.getRange("F"+i).getValue();
//     var description = transactionsSheet.getRange("G"+i).getValue();
//     var status = transactionsSheet.getRange("H"+i).getValue();
//     copyTransactionToPreviousTransactionsSheet(previousTransactionsSheet, prvsTransLastRow, id, date, outFlowAmount, inFlowAmount, category, account, description, status);
//     prvsTransLastRow++;
//     clearRowData(i);
//   }
// }

// function testRecurringTransactions(){
//   var debug_e = {
//     'day-of-month': 27 
//   };
//   // recurringTransactions(debug_e)
//   moveToPreviousTransactions(debug_e)
// }

// function test_onEdit() {
//   onEdit({
//     user : Session.getActiveUser().getEmail(),
//     source : SpreadsheetApp.getActiveSpreadsheet(),
//     range : SpreadsheetApp.getActiveSpreadsheet().getActiveCell(),
//     value : SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
//     authMode : "LIMITED"
//   });
// }

// function uuid() {
//   return Utilities.getUuid();
// }
