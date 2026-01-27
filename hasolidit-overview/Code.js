var props = PropertiesService.getUserProperties();
props.setProperty('netWorthSheet', 'מעקב שווי נקי');

// מעקב שווי נקי – filled on 1st of each month by trigger; sheet name = "מעקב שווי נקי YYYY"
var NET_WORTH_SHEET_BASE = 'מעקב שווי נקי';
var NET_WORTH_SHEET_FIRST_ROW = 6;   // first data row for credits
var NET_WORTH_DEBITS_FIRST_ROW = 23; // first data row for debits
var NET_WORTH_DATA_LAST_COL = 17;    // last column we write (B=2 .. Q=17)

// --- Net worth sheet helpers ---
function getNetWorthSheetForYear(year) {
  var base = props.getProperty('netWorthSheet') || NET_WORTH_SHEET_BASE;
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(base + ' ' + year);
}

/**
 * Create "מעקב שווי נקי YYYY" if missing: duplicate previous year, rename, clear monthly data.
 * Clears rows 6..22 (credits) and 23..end (debits), columns B–Q.
 * @param {number} year - e.g. 2026
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} the sheet, or null if template not found
 */
function createNetWorthSheetForYearIfMissing(year) {
  var base = props.getProperty('netWorthSheet') || NET_WORTH_SHEET_BASE;
  var name = base + ' ' + year;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  var templateName = base + ' ' + (year - 1);
  var template = ss.getSheetByName(templateName);
  if (!template) {
    Logger.log('Template sheet "' + templateName + '" not found. Create it first.');
    return null;
  }

  var newSheet = template.copyTo(ss);
  newSheet.setName(name);

  // Clear the cells we populate per month: credits block (rows 6–22), debits block (rows 23–end), cols B–Q
  var colStart = 2;
  newSheet.getRange(NET_WORTH_SHEET_FIRST_ROW, colStart, NET_WORTH_DEBITS_FIRST_ROW - 1, NET_WORTH_DATA_LAST_COL).clearContent();
  newSheet.getRange(NET_WORTH_DEBITS_FIRST_ROW, colStart, newSheet.getMaxRows(), NET_WORTH_DATA_LAST_COL).clearContent();

  Logger.log('Created and cleared sheet "' + name + '" from template "' + templateName + '"');
  return newSheet;
}

/**
 * Test: delete "מעקב שווי נקי 2026" if it exists, then create it from 2025 template and clear monthly data.
 * Run from Script Editor or via menu "Net worth functions" → "Create 2026 sheet (test)".
 */
function createAndSetup2026Sheet() {
  var base = props.getProperty('netWorthSheet') || NET_WORTH_SHEET_BASE;
  var name2026 = base + ' 2026';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(name2026);
  if (existing) {
    ss.deleteSheet(existing);
    Logger.log('Deleted existing sheet "' + name2026 + '"');
  }
  var sheet = createNetWorthSheetForYearIfMissing(2026);
  var ui = SpreadsheetApp.getUi();
  if (sheet) {
    if (ui) ui.toast('Sheet "' + name2026 + '" created and cleared.', 5);
    Logger.log('Sheet "' + name2026 + '" created and cleared.');
  } else {
    if (ui) ui.toast('Failed: template "' + base + ' 2025" not found.', 8);
    Logger.log('Failed: template "' + base + ' 2025" not found.');
  }
}

function test(){
  // var budgetSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1A2wZZwusoMeBLHQtbER6Qz3lasd1HoRSNZ3KmawRBUs/edit#gid=0");
  // var budgetSpreadSheetNetWorth = budgetSheet.getSheetByName("Net Worth Reports");
  // var mortgageLeft = mortgageValue(budgetSpreadSheetNetWorth);
  updateMortgageAndKupatGemelLeashkaaDebit()
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Functions')
      .addItem('fetch categories amounts', 'fetchCategoriesMonthlySums')
      .addItem('fetch categories amounts of previous year', 'fetchCategoriesMonthlySumsPreviousYear')
      .addSeparator()
      .addSubMenu(ui.createMenu('Net worth functions')
        .addItem('fetch clean worth amounts', 'fetchNetWorthMonthlySums')
        .addItem('fetch mortgage debt', 'updateMortgageAndKupatGemelLeashkaaDebit')
        .addItem('Create 2026 sheet (test)', 'createAndSetup2026Sheet'))
      .addToUi();
}

function fetchCategoriesMonthlySums() {
  var categoriesSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1A2wZZwusoMeBLHQtbER6Qz3lasd1HoRSNZ3KmawRBUs/edit#gid=1489150315");
  var categoryNames = categoriesSheet.getRangeByName("categories_report_names").getValues()
  var categorySums = categoriesSheet.getRangeByName("categories_report_sum").getValues()
  var currentYear = new Date().getFullYear();
  var calculationsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`Calculations${currentYear}`)
  calculationsSheet.getRange(1,1,categoryNames.length).setValues(categoryNames)
  calculationsSheet.getRange(1,2,categorySums.length, 14).setValues(categorySums)
}

function fetchCategoriesMonthlySumsPreviousYear() {
  var categoriesSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1A2wZZwusoMeBLHQtbER6Qz3lasd1HoRSNZ3KmawRBUs/edit#gid=1489150315");
  var categoryNames = categoriesSheet.getRangeByName("categories_report_names").getValues()
  var categorySums = categoriesSheet.getRangeByName("categories_report_sum").getValues()
  var currentYear = new Date().getFullYear() - 1;
  var calculationsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`Calculations${currentYear}`)
  calculationsSheet.getRange(1,1,categoryNames.length).setValues(categoryNames)
  calculationsSheet.getRange(1,2,categorySums.length, 14).setValues(categorySums)
}

/** Run on 1st of month. Fetches from portfolio + cashflow, appends one credits row and one debits row to "מעקב שווי נקי YYYY". Creates sheet from previous year if missing. */
function fetchNetWorthMonthlySums() {
  var investmentsSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1S8ploU9ZuZQGZ7B1AbH8lZk9SIFe-yqPdwOSxH1GH7U/edit#gid=997506131");
  var budgetSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1A2wZZwusoMeBLHQtbER6Qz3lasd1HoRSNZ3KmawRBUs/edit#gid=0");
  var reaHishtalmutAmount = investmentsSheet.getRangeByName('worth_hishtalmut').getValue();
  var reaPensionAmount = investmentsSheet.getRangeByName('worth_pension_estimation').getValue();
  var reaKupatGemel = investmentsSheet.getRangeByName('worth_kupat_gemel').getValue();
  var stocksWorth = Number(investmentsSheet.getRangeByName('worth_ib_stocks').getValue());
  var stocksAccountCashSheqel = Number(investmentsSheet.getRangeByName('trading_account_balance_sheqel').getValue());
  var stocksAccountCashUSDConverted = Number(investmentsSheet.getRangeByName('trading_account_balance_usd_converted').getValue());
  var savingsWorth = investmentsSheet.getRangeByName('worth_cash').getValue();
  var realEstateWorth = investmentsSheet.getRangeByName('worth_real_estate').getValue();
  var cryptoWorth = investmentsSheet.getRangeByName('worth_crypto').getValue();
  var otsarHayahalCheckingAccount = budgetSheet.getRangeByName('otsar_hahayal_checking_account').getValue();
  var oneZeroCheckingAccount = budgetSheet.getRangeByName('one_zero_checking_account').getValue();
  var creditCardsTotal = budgetSheet.getRangeByName('total_credit_cards').getValue();
  var currentYear = new Date().getFullYear();
  var netWorthSheet = getNetWorthSheetForYear(currentYear);
  if (!netWorthSheet) {
    netWorthSheet = createNetWorthSheetForYearIfMissing(currentYear);
  }
  if (!netWorthSheet) {
    Logger.log('Sheet מעקב שווי נקי ' + currentYear + ' not found and could not be created.');
    return;
  }

  // Credits: one row, columns B–Q (2–17). Order: checking, savings, 0, stocks+cash, crypto, hishtalmut, 0×4, [calculated col], pension+kupat, realEstate, 0×3.
  var checking = otsarHayahalCheckingAccount + oneZeroCheckingAccount;
  var stocksTotal = stocksWorth + stocksAccountCashSheqel + stocksAccountCashUSDConverted;
  var pensionKupat = reaPensionAmount + reaKupatGemel;
  var creditsRow = [
    checking,           // col B
    savingsWorth,       // col C
    0,                  // col D
    stocksTotal,        // col E
    cryptoWorth,        // col F
    reaHishtalmutAmount,// col G
    0, 0, 0, 0,         // cols H–K (kupat gemel, invest real estate, business, expensive material, other)
    0,                  // col L (calculated, leave 0 or formula handled elsewhere)
    pensionKupat,       // col M
    realEstateWorth,    // col N
    0, 0, 0             // cols O–Q (car, life insurance, collections)
  ];
  var netWorthLastRow = getNetWorthSheetLastRow(netWorthSheet);
  netWorthSheet.getRange(netWorthLastRow, 2, netWorthLastRow, 2 + creditsRow.length - 1).setValues([creditsRow]);

  // Debits: one row, columns B–J. We only set B (credit cards) and J.
  var netWorthDebitsLastRow = getNetWorthDebitsSheetLastRow(netWorthSheet);
  var debitsRow = [ -Number(creditCardsTotal), 0, 0, 0, 0, 0, 0, 0, 0 ];
  netWorthSheet.getRange(netWorthDebitsLastRow, 2, netWorthDebitsLastRow, 10).setValues([debitsRow]);
}

/** Run after fetchNetWorthMonthlySums. Backfills previous month's row with mortgage and Kupat Gemel from cashflow Net Worth Reports. */
function updateMortgageAndKupatGemelLeashkaaDebit() {
  var currentYear = new Date().getFullYear();
  var netWorthSheet = getNetWorthSheetForYear(currentYear);
  if (!netWorthSheet) return;
  var budgetSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1A2wZZwusoMeBLHQtbER6Qz3lasd1HoRSNZ3KmawRBUs/edit#gid=0");
  var budgetSpreadSheetNetWorth = budgetSheet.getSheetByName("Net Worth Reports");
  var netWorthDebitsLastRow = getNetWorthDebitsSheetLastRow(netWorthSheet)-1; // last month row was suppose to be updated
  Logger.log(`Last row: ${netWorthDebitsLastRow}`);
  var mortgageLeft = mortgageValue(budgetSpreadSheetNetWorth);
  netWorthSheet.getRange(netWorthDebitsLastRow, 10).setValue(`-${mortgageLeft}`);
  var analystValue = analystKupatGemelLastMonthValue(budgetSpreadSheetNetWorth);
  var meitavValue = meitavKupatGemelLastMonthValue(budgetSpreadSheetNetWorth);
  var netWorthLastRow = getNetWorthSheetLastRow(netWorthSheet)-1; // last month row was suppose to be updated
  netWorthSheet.getRange(netWorthLastRow, 8).setValue(analystValue + meitavValue);
}

// --- Last-row helpers (credits from row 6, debits from row 23) ---
function getNetWorthSheetLastRow(sheet) {
  var count = 0;
  var firstRow = NET_WORTH_SHEET_FIRST_ROW;
  for (var i = firstRow; !sheet.getRange('B' + i).isBlank(); i++) {
    count++;
  }
  return count + firstRow;
}

function getNetWorthDebitsSheetLastRow(sheet) {
  var count = 0;
  var firstRow = NET_WORTH_DEBITS_FIRST_ROW;
  for (var i = firstRow; !sheet.getRange('B' + i).isBlank(); i++) {
    count++;
  }
  return count + firstRow;
}

// --- Read from cashflow "Net Worth Reports" (for backfill) ---
function mortgageValue(sheet) {
  var textFinder = sheet.createTextFinder("Mortgage");
  var allResult = textFinder.findAll();
  var mortgageRow = allResult[allResult.length-1].getRow();
  return sheet.getRange(`C${mortgageRow}`).getValue();
}

function analystKupatGemelLastMonthValue(sheet) {
  var textFinder = sheet.createTextFinder("Analyst Kupat Gemel Le'ashkaa");
  var allResult = textFinder.findAll();
  var row = allResult[allResult.length-1].getRow();
  return sheet.getRange(`C${row}`).getValue();
}

function meitavKupatGemelLastMonthValue(sheet) {
  var textFinder = sheet.createTextFinder("Meitav Kupat Gemel Le'ashkaa");
  var allResult = textFinder.findAll();
  var row = allResult[allResult.length-1].getRow();
  return sheet.getRange(`C${row}`).getValue();
}

function sheetName(idx) {
  if (!idx)
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  else {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var idx = parseInt(idx);
    if (isNaN(idx) || idx < 1 || sheets.length < idx)
      throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
    return sheets[idx-1].getName();
  }
}

function yearInSheetName(idx) {
  var name = sheetName(idx)
  var numberPattern = /\d+/g;
  value = name.match( numberPattern ).join([]);
  return value
}
