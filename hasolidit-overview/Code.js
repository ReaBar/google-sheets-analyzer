var props = PropertiesService.getUserProperties();
props.setProperty('netWorthSheet', 'מעקב שווי נקי');

// מעקב שווי נקי – filled on 1st of each month by trigger; sheet name = "מעקב שווי נקי YYYY"
var NET_WORTH_SHEET_BASE = 'מעקב שווי נקי';
var NET_WORTH_SHEET_FIRST_ROW = 6;        // sub-header row for credits (do not clear)
var NET_WORTH_DEBITS_FIRST_ROW = 23;      // sub-header row for debits (do not clear)
var NET_WORTH_CREDITS_FIRST_DATA_ROW = 7; // first data row for credits
var NET_WORTH_CREDITS_LAST_DATA_ROW = 22; // last data row for credits block (before debits header)
var NET_WORTH_DEBITS_FIRST_DATA_ROW = 24; // first data row for debits
var NET_WORTH_DEBITS_LAST_DATA_ROW = 35;  // last data row for debits block (before analysis section)
var NET_WORTH_DATA_LAST_COL = 17;         // last column we write (B=2 .. Q=17)
var NET_WORTH_FORMULA_COLUMN_L = 12; // column L: formulas live here; do not clear or overwrite

// --- Formula definitions (single source of truth; edit here and clasp push; script writes these into the sheet) ---
// When do they run? The script only writes formula strings into cells. Google Sheets then evaluates them
// on normal recalc (on open, on edit, when dependencies change). No onEdit needed.
// Use {row} as placeholder; it is replaced with the actual row number when writing.
var NET_WORTH_FORMULAS = {
  // Credits block, column L: total of cols B–K for this row
  creditsTotalL: '=SUM(B{row}:K{row})',
  // Debits block, column L: total of cols B–I for this row (adjust if your layout differs)
  debitsTotalL: '=SUM(B{row}:I{row})'
};

function getNetWorthFormula(formulaKey, row) {
  var template = NET_WORTH_FORMULAS[formulaKey];
  return template ? template.replace(/\{row\}/g, String(row)) : '';
}

/**
 * Apply NET_WORTH_FORMULAS to column L for all data rows (credits 7–22, debits 24–last used).
 * Call after clearing when creating a new year sheet so formulas come from the script.
 */
function applyNetWorthFormulasToColumnL(sheet) {
  var colL = NET_WORTH_FORMULA_COLUMN_L;
  var r;
  for (r = NET_WORTH_CREDITS_FIRST_DATA_ROW; r <= NET_WORTH_CREDITS_LAST_DATA_ROW; r++) {
    var f = getNetWorthFormula('creditsTotalL', r);
    if (f) sheet.getRange(r, colL).setFormula(f);
  }
  for (r = NET_WORTH_DEBITS_FIRST_DATA_ROW; r <= NET_WORTH_DEBITS_LAST_DATA_ROW; r++) {
    var g = getNetWorthFormula('debitsTotalL', r);
    if (g) sheet.getRange(r, colL).setFormula(g);
  }
}

/**
 * Clear monthly data ranges while preserving any existing formulas.
 * Credits: rows 7–22, Debits: rows 24–35, columns B–Q. Rows 6 and 23 are sub-headers.
 */
function clearNetWorthDataPreservingFormulas(sheet) {
  var firstCol = 2; // B
  var lastCol = NET_WORTH_DATA_LAST_COL; // Q

  // Helper to clear non-formula cells in a row range, using batched operations:
  // - Fetch all formulas for the block once with getFormulas()
  // - For each row, clear only contiguous runs of cells that do NOT contain
  //   *protected* formulas (we still clear formulas in data columns like B7).
  // This avoids per-cell getRange/getFormula calls and is much faster.
  //
  // protectedFormulaCols is an array of 1-based column numbers whose formulas
  // we want to keep (e.g. credits totals in column M, debits totals in H and O).
  function clearBlock(fromRow, toRow, protectedFormulaCols) {
    var numRows = toRow - fromRow + 1;
    var numCols = lastCol - firstCol + 1;
    if (numRows <= 0 || numCols <= 0) return;

    var formulas = sheet
      .getRange(fromRow, firstCol, numRows, numCols)
      .getFormulas();

    for (var i = 0; i < numRows; i++) {
      var rowIndex = fromRow + i;
      var inRun = false;
      var runStart = 0;

      for (var j = 0; j < numCols; j++) {
        var formulaText = (formulas[i][j] || '').toString().trim();
        var absoluteCol = firstCol + j;
        var isProtectedFormula =
          formulaText !== '' &&
          protectedFormulaCols &&
          protectedFormulaCols.indexOf(absoluteCol) !== -1;

        // Start of a run of non-formula cells
        if (!isProtectedFormula && !inRun) {
          inRun = true;
          runStart = j;
        }

        var atLastCol = j === numCols - 1;
        // End of a run: either we hit a formula or the last column
        if ((isProtectedFormula || atLastCol) && inRun) {
          var runEnd = isProtectedFormula ? j - 1 : j;
          if (runEnd >= runStart) {
            var colStart = firstCol + runStart;
            var width = runEnd - runStart + 1;
            sheet.getRange(rowIndex, colStart, 1, width).clearContent();
          }
          inRun = false;
        }
      }
    }
  }

  // Credits: keep formulas only in column M (13) within B–Q
  clearBlock(NET_WORTH_CREDITS_FIRST_DATA_ROW, NET_WORTH_CREDITS_LAST_DATA_ROW, [13]);
  // Debits: keep formulas only in columns H (8) and O (15) within B–Q
  clearBlock(NET_WORTH_DEBITS_FIRST_DATA_ROW, NET_WORTH_DEBITS_LAST_DATA_ROW, [8, 15]);
}

// --- Net worth sheet helpers ---
function getNetWorthSheetForYear(year) {
  var base = props.getProperty('netWorthSheet') || NET_WORTH_SHEET_BASE;
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(base + ' ' + year);
}

/**
 * Create "מעקב שווי נקי YYYY" if missing: duplicate previous year, rename, clear monthly data.
 * Clears credits (7–22) and debits (24–end) in B–K and M–Q only; column L has formulas and is not cleared.
 * Rows 6 and 23 are sub-headers and are left as-is.
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

  // Clear monthly data only, preserving any cells that currently contain formulas.
  // Credits: rows 7–22, Debits: rows 24–35, columns B–Q. Rows 6 and 23 are sub-headers.
  clearNetWorthDataPreservingFormulas(newSheet);

  // Write formula column L from script (single source of truth)
  applyNetWorthFormulasToColumnL(newSheet);

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
  if (sheet) {
    ss.toast('Sheet "' + name2026 + '" created and cleared.', 5);
    Logger.log('Sheet "' + name2026 + '" created and cleared.');
  } else {
    ss.toast('Failed: template "' + base + ' 2025" not found.', 8);
    Logger.log('Failed: template "' + base + ' 2025" not found.');
  }
}

/**
 * Export all formulas from every sheet to a dedicated sheet "_FormulaExport".
 * Run from menu "Export formulas to sheet". Then download that sheet as CSV or copy into the repo
 * so formulas can be reviewed or migrated into NET_WORTH_FORMULAS in Code.js.
 */
function exportFormulasToSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var exportName = '_FormulaExport';
  var exportSheet = ss.getSheetByName(exportName);
  if (!exportSheet) {
    exportSheet = ss.insertSheet(exportName);
  }
  exportSheet.clear();

  var out = [['Sheet', 'Cell', 'Formula']];
  var sheets = ss.getSheets();

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var range = sheet.getDataRange();
    if (!range) continue;
    var formulas = range.getFormulas();
    var rowOffset = range.getRow();
    var colOffset = range.getColumn();

    for (var i = 0; i < formulas.length; i++) {
      for (var j = 0; j < (formulas[i] || []).length; j++) {
        var f = (formulas[i][j] || '').toString().trim();
        if (f) {
          // Strip leading '=' so export shows the formula text, not a live formula that re-evaluates.
          var formulaText = f.charAt(0) === '=' ? f.substring(1) : f;
          var a1 = colToLetter(colOffset + j) + (rowOffset + i);
          out.push([sheet.getName(), a1, formulaText]);
        }
      }
    }
  }

  if (out.length <= 1) {
    exportSheet.getRange(1, 1).setValue('No formulas found.');
  } else {
    exportSheet.getRange(1, 1, out.length, 3).setValues(out);
    exportSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  ss.toast('Formulas exported to sheet "' + exportName + '". Download as CSV or copy to share.', 8);
}

function colToLetter(col) {
  var letter = '';
  var n = col;
  while (n > 0) {
    n--;
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26);
  }
  return letter;
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
      .addSeparator()
      .addItem('Export formulas to sheet', 'exportFormulasToSheet')
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

/** Run on 1st of month (or manually). Fetches from portfolio + cashflow, fills the row for the current month
 *  in "מעקב שווי נקי YYYY" (January → row 7 credits / 24 debits, February → row 8 / 25, etc.).
 *  Only empty cells are populated so existing values are not overwritten on repeated runs.
 *  Creates the yearly sheet from the previous year if missing.
 */
function fetchNetWorthMonthlySums() {
  var investmentsSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1S8ploU9ZuZQGZ7B1AbH8lZk9SIFe-yqPdwOSxH1GH7U/edit#gid=997506131");
  var budgetSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1A2wZZwusoMeBLHQtbER6Qz3lasd1HoRSNZ3KmawRBUs/edit#gid=0");
  var reaHishtalmutAmount = investmentsSheet.getRangeByName('worth_hishtalmut').getValue();
  var reaPensionAmount = investmentsSheet.getRangeByName('worth_pension_estimation').getValue();
  var reaKupatGemel = investmentsSheet.getRangeByName('worth_kupat_gemel').getValue();
  var stocksWorth = Number(investmentsSheet.getRangeByName('worth_ib_stocks').getValue());
  var stocksAccountCashSheqel = Number(investmentsSheet.getRangeByName('trading_account_balance_sheqel').getValue());
  var stocksAccountCashUSDConverted = Number(investmentsSheet.getRangeByName('trading_account_balance_usd_converted').getValue());
  // Savings (column C) now comes from cashflow's one_zero_savings_balance named range, not from investments worth_cash.
  var savingsWorth = budgetSheet.getRangeByName('one_zero_savings_balance').getValue();
  var realEstateWorth = investmentsSheet.getRangeByName('worth_real_estate').getValue();
  var cryptoWorth = investmentsSheet.getRangeByName('worth_crypto').getValue();
  var otsarHayahalCheckingAccount = budgetSheet.getRangeByName('otsar_hahayal_checking_account').getValue();
  var oneZeroCheckingAccount = budgetSheet.getRangeByName('one_zero_checking_account').getValue();
  var creditCardsTotal = budgetSheet.getRangeByName('total_credit_cards').getValue();

  var now = new Date();
  var currentYear = now.getFullYear();
  var monthIndex = now.getMonth(); // 0 = Jan, 1 = Feb, ...

  var netWorthSheet = getNetWorthSheetForYear(currentYear);
  if (!netWorthSheet) {
    netWorthSheet = createNetWorthSheetForYearIfMissing(currentYear);
  }
  if (!netWorthSheet) {
    Logger.log('Sheet מעקב שווי נקי ' + currentYear + ' not found and could not be created.');
    return;
  }

  // Determine fixed rows for this month: credits (row 7–18), debits (row 24–35)
  var creditsRow = NET_WORTH_CREDITS_FIRST_DATA_ROW + monthIndex;   // Jan=7, Feb=8, ...
  var debitsRow = NET_WORTH_DEBITS_FIRST_DATA_ROW + monthIndex;     // Jan=24, Feb=25, ...

  // Credits: one row, columns B–K and N–Q. Columns L and M have formulas and are not written.
  var checking = otsarHayahalCheckingAccount + oneZeroCheckingAccount;
  var stocksTotal = stocksWorth + stocksAccountCashSheqel + stocksAccountCashUSDConverted;
  var pensionKupat = reaPensionAmount + reaKupatGemel;
  var creditsBtoK = [
    checking, savingsWorth, 0, stocksTotal, cryptoWorth, reaHishtalmutAmount,
    0, 0, 0, 0  // cols H–K
  ];
  var creditsNtoQ = [ pensionKupat, realEstateWorth, 0, 0, 0 ];  // cols N–Q

  // Write credits for this month, but only into empty cells.
  var rangeBtoK = netWorthSheet.getRange(creditsRow, 2, 1, creditsBtoK.length); // B–K
  var existingBtoK = rangeBtoK.getValues()[0];
  for (var i = 0; i < creditsBtoK.length; i++) {
    if (existingBtoK[i] === '' || existingBtoK[i] === null) {
      existingBtoK[i] = creditsBtoK[i];
    }
  }
  rangeBtoK.setValues([existingBtoK]);

  var rangeNtoQ = netWorthSheet.getRange(creditsRow, 14, 1, creditsNtoQ.length); // N–Q
  var existingNtoQ = rangeNtoQ.getValues()[0];
  for (var j = 0; j < creditsNtoQ.length; j++) {
    if (existingNtoQ[j] === '' || existingNtoQ[j] === null) {
      existingNtoQ[j] = creditsNtoQ[j];
    }
  }
  rangeNtoQ.setValues([existingNtoQ]);

  // Debits: one row, columns B–J. Only empty cells are updated; formulas in H/O remain untouched.
  var debitsValues = [ -Number(creditCardsTotal), 0, 0, 0, 0, 0, 0, 0, 0 ]; // B–J (9 columns)
  var debitsRange = netWorthSheet.getRange(debitsRow, 2, 1, debitsValues.length);
  var existingDebits = debitsRange.getValues()[0];
  for (var k = 0; k < debitsValues.length; k++) {
    if (existingDebits[k] === '' || existingDebits[k] === null) {
      existingDebits[k] = debitsValues[k];
    }
  }
  debitsRange.setValues([existingDebits]);
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

// --- Last-row helpers (credits from row 7, debits from row 24; rows 6 and 23 are sub-headers) ---
function getNetWorthSheetLastRow(sheet) {
  var count = 0;
  var firstRow = NET_WORTH_CREDITS_FIRST_DATA_ROW;
  for (var i = firstRow; !sheet.getRange('B' + i).isBlank(); i++) {
    count++;
  }
  return count + firstRow;
}

function getNetWorthDebitsSheetLastRow(sheet) {
  var count = 0;
  var firstRow = NET_WORTH_DEBITS_FIRST_DATA_ROW;
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
