/**
 * Expense Tracker — Splitwise-style for Google Sheets
 * V4 — Clean rebuild: single source of truth + 100% live formulas
 * Developed by Athreyas using Claude (Anthropic AI)
 *
 * ARCHITECTURE (why this version is reliable):
 *  - "Expenses" sheet is the ONLY source of truth (what you type lives here).
 *  - "Ledger" (hidden) is auto-derived from Expenses: one row per person per
 *    transaction, holding what they Paid and what their Share was.
 *  - Dashboard + Balances show ONLY live formulas that read the Ledger/Expenses.
 *    The script never writes a balance number, so nothing can fall out of sync.
 *  - Adding a user / expense / settlement just appends data, then rebuilds the
 *    Ledger from scratch. "Refresh Balances" does the same — fully self-healing.
 */

var CONFIG = {
  colors: {
    headerBg: '#1a73e8',
    headerText: '#ffffff',
    primaryBg: '#e8f0fe',
    secondaryBg: '#f8f9fa',
    userColors: ['#fce4ec', '#e8f5e9', '#fff3e0', '#e1f5fe', '#f3e5f5', '#e0f2f1', '#fff9c4', '#ede7f6'],
    positiveGreen: '#c8e6c9',
    negativeRed: '#ffcdd2',
    borderColor: '#dadce0'
  },
  sheets: {
    about: 'About',
    dashboard: 'Dashboard',
    balances: 'Balances',
    expenses: 'Expenses',
    users: 'Master Users',
    ledger: 'Ledger',
    archives: 'Archives'
  },
  fonts: {
    header: 'Product Sans',
    body: 'Google Sans',
    mono: 'Roboto Mono'
  }
};

/* ============================================================
 *  MENU + DIALOGS
 * ============================================================ */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Expense Manager')
    .addItem('🔧 Setup (Admin Only)', 'showSetupDialog')
    .addItem('➕ Add New User', 'showAddUserDialog')
    .addSeparator()
    .addItem('➕ Add Expense', 'showAddExpenseDialog')
    .addItem('💳 Add Settlement', 'showAddSettlementDialog')
    .addSeparator()
    .addItem('🔄 Refresh Balances', 'refreshBalancesWithUI')
    .addItem('📦 Archive Old Expenses', 'showArchiveDialog')
    .addSeparator()
    .addItem('❓ Help', 'showHelp')
    .addToUi();

  // Land on the Dashboard each time the file opens.
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dash = ss.getSheetByName(CONFIG.sheets.dashboard);
    if (dash) ss.setActiveSheet(dash);
  } catch (e) { /* ignore */ }
}

function showSetupDialog() {
  var ui = SpreadsheetApp.getUi();
  if (isSetupComplete()) {
    var response = ui.alert(
      '⚠️ Warning: Setup Already Complete',
      'Running setup again will DELETE ALL existing data including expenses, settlements, and balances.\n\nThis action CANNOT be undone!\n\nContinue?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    var confirmEmail = ui.prompt('Admin Verification Required', 'Enter your email to confirm:', ui.ButtonSet.OK_CANCEL);
    if (confirmEmail.getSelectedButton() !== ui.Button.OK || !confirmEmail.getResponseText().trim()) {
      ui.alert('Setup cancelled.');
      return;
    }
    PropertiesService.getScriptProperties().setProperty('ADMIN_EMAIL', confirmEmail.getResponseText().trim());
  }
  var html = HtmlService.createHtmlOutputFromFile('SetupDialog').setWidth(500).setHeight(450);
  ui.showModalDialog(html, '🔧 Setup Expense Tracker');
}

function showAddUserDialog() {
  if (!isSetupComplete()) { SpreadsheetApp.getUi().alert('Please complete setup first.'); return; }
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('➕ Add New User', 'Enter the name of the new user:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var userName = response.getResponseText().trim();
    if (userName) ui.alert(addNewUser(userName).message);
  }
}

function showAddExpenseDialog() {
  if (!isSetupComplete()) { SpreadsheetApp.getUi().alert('Please complete setup first.'); return; }
  var html = HtmlService.createHtmlOutputFromFile('AddExpenseDialog').setWidth(600).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Add New Expense');
}

function showAddSettlementDialog() {
  if (!isSetupComplete()) { SpreadsheetApp.getUi().alert('Please complete setup first.'); return; }
  var html = HtmlService.createHtmlOutputFromFile('AddSettlementDialog').setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '💳 Add Settlement');
}

function showArchiveDialog() {
  if (!isSetupComplete()) { SpreadsheetApp.getUi().alert('Please complete setup first.'); return; }
  var html = HtmlService.createHtmlOutputFromFile('ArchiveDialog').setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '📦 Archive Expenses');
}

function showHelp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aboutSheet = ss.getSheetByName(CONFIG.sheets.about);
  if (aboutSheet) ss.setActiveSheet(aboutSheet);
  SpreadsheetApp.getUi().alert('❓ Help', 'Please refer to the "About" sheet for detailed instructions.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function isSetupComplete() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.users) !== null;
}

/* ============================================================
 *  SETUP
 * ============================================================ */

function processSetup(users) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Remove all but one sheet so we can rebuild cleanly.
    var sheets = ss.getSheets();
    for (var i = sheets.length - 1; i >= 1; i--) ss.deleteSheet(sheets[i]);
    if (ss.getSheets().length === 0) ss.insertSheet('Temp');

    var userArray = users.split('\n').map(function (u) { return u.trim(); }).filter(function (u) { return u !== ''; });

    setupAboutSheet();
    createMasterUsersSheet(userArray);
    createExpensesSheet(userArray);
    createLedgerSheet();
    createArchivesSheet();
    createDashboardSheet(userArray);

    var tempSheet = ss.getSheetByName('Temp');
    if (tempSheet) ss.deleteSheet(tempSheet);

    reorderSheets();
    rebuildLedger(); // empty to start, but establishes a clean state

    ss.setActiveSheet(ss.getSheetByName(CONFIG.sheets.dashboard));
    return { success: true, message: '✅ Setup completed successfully!' };
  } catch (error) {
    Logger.log('Setup error: ' + error.toString());
    return { success: false, message: '❌ Error: ' + error.toString() };
  }
}

function reorderSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var order = [CONFIG.sheets.about, CONFIG.sheets.dashboard, CONFIG.sheets.expenses];
  for (var i = 0; i < order.length; i++) {
    var sheet = ss.getSheetByName(order[i]);
    if (sheet && !sheet.isSheetHidden()) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    }
  }
}

/* ============================================================
 *  SHEET BUILDERS
 * ============================================================ */

function setupAboutSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var existingAbout = ss.getSheetByName(CONFIG.sheets.about);
  if (existingAbout && existingAbout.getSheetId() !== sheet.getSheetId()) ss.deleteSheet(existingAbout);

  sheet.setName(CONFIG.sheets.about);
  sheet.clear();
  sheet.setColumnWidth(1, 155);
  sheet.setColumnWidth(2, 645);

  var headerRange = sheet.getRange('A1:B1').merge();
  headerRange.setValue('💰 EXPENSE TRACKER — QUICK START GUIDE')
    .setFontFamily(CONFIG.fonts.header).setFontSize(22).setFontWeight('bold')
    .setBackground(CONFIG.colors.headerBg).setFontColor(CONFIG.colors.headerText)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 56);

  var r = 2;
  sheet.getRange(r, 1, 1, 2).merge();
  sheet.getRange(r, 1).setValue('📝 Developed by Athreyas using Claude (Anthropic AI)')
    .setFontFamily(CONFIG.fonts.body).setFontStyle('italic').setFontSize(10)
    .setBackground('#f1f3f4').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(r, 26); r++;

  sheet.getRange(r, 1, 1, 2).merge();
  sheet.getRange(r, 1).setFormula('=HYPERLINK("https://github.com/Athreyas/googlesheets_expense_tracker","🔗 View this project on GitHub")')
    .setFontFamily(CONFIG.fonts.body).setFontSize(10).setFontColor('#1a73e8')
    .setBackground('#f1f3f4').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(r, 26); r++;

  sheet.getRange(r, 1, 1, 2).merge();
  sheet.getRange(r, 1).setValue('Welcome! This tool helps you track shared expenses with roommates, friends, or family. Split bills fairly and see who owes whom at a glance.')
    .setFontFamily(CONFIG.fonts.body).setFontSize(12).setBackground('#e8f0fe')
    .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 50); r += 2;

  r = aboutSection(sheet, r, '✨ KEY FEATURES', '#34a853', [
    ['📊', 'Live dashboard: month stats, per-person balances, and suggested settlements'],
    ['💳', 'Multiple split types: Equal, Exact amounts, Percentages, or Shares'],
    ['⚖️', 'Net balance per person showing exactly who is up and who is down'],
    ['🔄', 'Automatic calculations - balances update the moment you add data'],
    ['📦', 'Archive old expenses to keep current data clean'],
    ['➕', 'Add new users anytime without losing data']
  ], '#f8f9fa', 32, 11, 16);

  r++;
  r = aboutSection(sheet, r, '🚀 QUICK START (3 SIMPLE STEPS)', '#ea4335', [
    ['1️⃣', 'ADD AN EXPENSE\nMenu → Add Expense → Fill details → Choose split type → Submit\nExample: "Groceries $120, paid by John, split equally among everyone"'],
    ['2️⃣', 'CHECK THE DASHBOARD\nThe Dashboard shows spending, net balances, and suggested settlements all in one place\nGreen = owed to them | Red = they owe'],
    ['3️⃣', 'RECORD SETTLEMENT\nMenu → Add Settlement → Select who paid whom → Enter amount\nBalances update automatically!']
  ], '#fff3cd', 65, 11, 24);

  r++;
  r = aboutSection(sheet, r, '💡 UNDERSTANDING SPLIT TYPES', '#fbbc04', [
    ['EQUAL', '$120 ÷ 3 people = $40 each\nUse when: Everyone pays the same amount'],
    ['EXACT', 'Alice: $50, Bob: $40, Carol: $30 (Total: $120)\nUse when: Different people owe different specific amounts'],
    ['PERCENTAGE', 'Alice: 50% ($60), Bob: 30% ($36), Carol: 20% ($24)\nUse when: Splitting by agreed percentages (must total 100%)'],
    ['SHARES', 'Alice: 2, Bob: 1, Carol: 1 (ratio 2:1:1)\nUse when: Splitting by ratio (e.g., roommates by room size)']
  ], '#e0f2f1', 55, 10, 11);

  r++;
  r = aboutSection(sheet, r, '📱 SHEET NAVIGATION', '#9c27b0', [
    ['📊 DASHBOARD', 'Month summary, spending & net balance per person, and suggested settlements'],
    ['💳 EXPENSES', 'All transaction history | See split details for every entry']
  ], '#f3e5f5', 38, 10, 11);

  r++;
  // Important notes (full-width rows)
  sheet.getRange(r, 1, 1, 2).merge();
  sheet.getRange(r, 1).setValue('⚠️ IMPORTANT NOTES')
    .setFontFamily(CONFIG.fonts.header).setFontSize(14).setFontWeight('bold')
    .setBackground('#ff6f00').setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 32); r++;
  var notes = [
    '• All sheets are protected - use the menu options to make changes',
    '• Balances update automatically after adding expenses or settlements',
    '• Use "Refresh Balances" to fully recalculate everything from scratch',
    '• Setup will DELETE all data - use with extreme caution',
    '• Archive expenses periodically to keep the system fast'
  ];
  for (var n = 0; n < notes.length; n++) {
    sheet.getRange(r, 1, 1, 2).merge();
    sheet.getRange(r, 1).setValue(notes[n])
      .setFontFamily(CONFIG.fonts.body).setFontSize(10).setWrap(true).setBackground('#fff3e0')
      .setBorder(true, true, true, true, false, false, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(r, 28); r++;
  }

  r++;
  sheet.getRange(r, 1, 1, 2).merge();
  sheet.getRange(r, 1).setValue('Need help? Click: 💰 Expense Manager → ❓ Help')
    .setFontFamily(CONFIG.fonts.body).setFontStyle('italic').setFontSize(11)
    .setBackground('#cfd8dc').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(r, 35);

  if (sheet.getMaxColumns() > 2) sheet.hideColumns(3, sheet.getMaxColumns() - 2);
  lockSheet(sheet, 'About');
  sheet.setFrozenRows(1);
}

/** Helper that renders a two-column titled section on the About sheet. */
function aboutSection(sheet, startRow, title, titleBg, rows, bodyBg, rowH, bodyFont, iconFont) {
  var r = startRow;
  sheet.getRange(r, 1, 1, 2).merge();
  sheet.getRange(r, 1).setValue(title)
    .setFontFamily(CONFIG.fonts.header).setFontSize(14).setFontWeight('bold')
    .setBackground(titleBg).setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 32); r++;

  for (var i = 0; i < rows.length; i++) {
    sheet.getRange(r, 1).setValue(rows[i][0])
      .setFontFamily(CONFIG.fonts.header).setFontWeight('bold').setFontSize(iconFont)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(bodyBg);
    sheet.getRange(r, 2).setValue(rows[i][1])
      .setFontFamily(CONFIG.fonts.body).setFontSize(bodyFont).setWrap(true)
      .setVerticalAlignment('middle').setBackground(bodyBg);
    sheet.setRowHeight(r, rowH);
    sheet.getRange(r, 1, 1, 2).setBorder(true, true, true, true, false, false, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  }
  return r;
}

function createMasterUsersSheet(userArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(CONFIG.sheets.users);
  if (existing) ss.deleteSheet(existing);

  var sheet = ss.insertSheet(CONFIG.sheets.users);
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 100);

  sheet.getRange('A1:C1').merge().setValue('👥 MASTER USERS')
    .setFontFamily(CONFIG.fonts.header).setFontSize(24).setFontWeight('bold')
    .setBackground(CONFIG.colors.headerBg).setFontColor(CONFIG.colors.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 60);

  sheet.getRange(2, 1, 1, 3).setValues([['ID', 'Name', 'Color']])
    .setFontFamily(CONFIG.fonts.header).setFontWeight('bold').setFontSize(11)
    .setBackground('#f1f3f4').setFontColor('#202124').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 40);

  for (var i = 0; i < userArray.length; i++) {
    var color = CONFIG.colors.userColors[i % CONFIG.colors.userColors.length];
    sheet.getRange(i + 3, 1).setValue(i + 1).setFontFamily(CONFIG.fonts.body).setHorizontalAlignment('center');
    sheet.getRange(i + 3, 2).setValue(userArray[i]).setFontFamily(CONFIG.fonts.body).setFontWeight('bold');
    sheet.getRange(i + 3, 3).setValue(color).setBackground(color).setFontFamily(CONFIG.fonts.mono).setFontSize(9);
    sheet.setRowHeight(i + 3, 32);
  }
  if (userArray.length > 0) {
    sheet.getRange(3, 1, userArray.length, 3).setBorder(true, true, true, true, true, true, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);
  }
  lockSheet(sheet, 'Master Users');
  sheet.setFrozenRows(2);
  sheet.hideSheet();
}

function createExpensesSheet(userArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(CONFIG.sheets.expenses);
  if (existing) ss.deleteSheet(existing);

  var sheet = ss.insertSheet(CONFIG.sheets.expenses);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 350);
  sheet.setColumnWidth(7, 140);
  sheet.setColumnWidth(8, 110);

  sheet.getRange('A1:H1').merge().setValue('💳 EXPENSE TRACKER')
    .setFontFamily(CONFIG.fonts.header).setFontSize(26).setFontWeight('bold')
    .setBackground(CONFIG.colors.headerBg).setFontColor(CONFIG.colors.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 70);

  sheet.getRange(2, 1, 1, 8).setValues([['Date', 'Description', 'Amount', 'Paid By', 'Split Type', 'Split Details', 'Category', 'Type']])
    .setFontFamily(CONFIG.fonts.header).setFontWeight('bold').setFontSize(11)
    .setBackground('#f1f3f4').setFontColor('#202124')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(2, 40);

  sheet.getRange('A3:A1000').setNumberFormat('mmm dd, yyyy').setHorizontalAlignment('center').setFontFamily(CONFIG.fonts.mono);
  sheet.getRange('C3:C1000').setNumberFormat('$#,##0.00').setHorizontalAlignment('right').setFontWeight('bold');
  sheet.getRange('E3:E1000').setHorizontalAlignment('center');
  sheet.getRange('H3:H1000').setHorizontalAlignment('center').setFontWeight('bold');

  if (sheet.getMaxColumns() > 8) sheet.hideColumns(9, sheet.getMaxColumns() - 8);
  applyExpensesConditionalFormatting(sheet, getNamesFromArray(userArray));
  lockSheet(sheet, 'Expenses');
  sheet.setFrozenRows(2);
}

function createLedgerSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(CONFIG.sheets.ledger);
  if (existing) ss.deleteSheet(existing);

  var sheet = ss.insertSheet(CONFIG.sheets.ledger);
  sheet.getRange('A1:F1').merge().setValue('🧮 LEDGER (auto-generated — do not edit)')
    .setFontFamily(CONFIG.fonts.header).setFontSize(16).setFontWeight('bold')
    .setBackground('#455a64').setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);

  sheet.getRange(2, 1, 1, 6).setValues([['Date', 'Person', 'Paid', 'Share', 'Type', 'Description']])
    .setFontFamily(CONFIG.fonts.header).setFontWeight('bold').setFontSize(11)
    .setBackground('#cfd8dc').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 32);

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 320);
  sheet.getRange('A3:A').setNumberFormat('mmm dd, yyyy');
  sheet.getRange('C3:D').setNumberFormat('$#,##0.00');

  lockSheet(sheet, 'Ledger');
  sheet.setFrozenRows(2);
  sheet.hideSheet();
}

function createArchivesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(CONFIG.sheets.archives);
  if (existing) ss.deleteSheet(existing);

  var sheet = ss.insertSheet(CONFIG.sheets.archives);
  for (var i = 1; i <= 9; i++) sheet.setColumnWidth(i, i === 2 ? 280 : i === 6 ? 350 : 110);

  sheet.getRange('A1:I1').merge().setValue('📦 ARCHIVED EXPENSES')
    .setFontFamily(CONFIG.fonts.header).setFontSize(24).setFontWeight('bold')
    .setBackground('#6c757d').setFontColor(CONFIG.colors.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 60);

  sheet.getRange(2, 1, 1, 9).setValues([['Date', 'Description', 'Amount', 'Paid By', 'Split Type', 'Split Details', 'Category', 'Type', 'Archived On']])
    .setFontFamily(CONFIG.fonts.header).setFontWeight('bold').setFontSize(11)
    .setBackground('#adb5bd').setFontColor('#202124').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 40);

  sheet.getRange('A3:A').setNumberFormat('mmm dd, yyyy').setHorizontalAlignment('center').setFontFamily(CONFIG.fonts.mono);
  sheet.getRange('C3:C').setNumberFormat('$#,##0.00').setHorizontalAlignment('right').setFontWeight('bold');
  sheet.getRange('I3:I').setNumberFormat('mmm dd, yyyy').setHorizontalAlignment('center').setFontFamily(CONFIG.fonts.mono);

  if (sheet.getMaxColumns() > 9) sheet.hideColumns(10, sheet.getMaxColumns() - 9);
  lockSheet(sheet, 'Archives');
  sheet.setFrozenRows(2);
  sheet.hideSheet();
}

/**
 * DASHBOARD — the single overview sheet (the old Balances tab is merged in).
 *   Top: a compact KPI band (four cards).
 *   Left region (A–E): spending + net balance per person — all live formulas.
 *   Right region (F–H): suggested settlements — regenerated on every rebuild.
 * Both tables are the last thing in their columns, so adding a user just
 * appends a row with nothing below to overwrite.
 */
var DASH = {
  personHeaderRow: 10, // "Person | Total Spent | # Transactions | Net Balance | Status"
  personDataRow: 11,
  sugHeaderRow: 10,    // "Who Pays | Who Receives | Amount"
  sugDataRow: 11,
  sugFirstCol: 6       // column F
};

function createDashboardSheet(userArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(CONFIG.sheets.dashboard);
  if (existing) ss.deleteSheet(existing);

  var sheet = ss.insertSheet(CONFIG.sheets.dashboard);
  sheet.setColumnWidth(1, 180); // Person
  sheet.setColumnWidth(2, 120); // Total Spent
  sheet.setColumnWidth(3, 120); // # Transactions
  sheet.setColumnWidth(4, 130); // Net Balance
  sheet.setColumnWidth(5, 150); // Status
  sheet.setColumnWidth(6, 150); // Who Pays
  sheet.setColumnWidth(7, 160); // Who Receives
  sheet.setColumnWidth(8, 120); // Amount

  // Title
  sheet.getRange('A1:H1').merge().setValue('📊 EXPENSE DASHBOARD')
    .setFontFamily(CONFIG.fonts.header).setFontSize(22).setFontWeight('bold')
    .setBackground(CONFIG.colors.headerBg).setFontColor(CONFIG.colors.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 54);

  // Live date
  sheet.getRange('A2:H2').merge().setFormula('=TEXT(TODAY(),"dddd, MMMM dd, yyyy")')
    .setFontFamily(CONFIG.fonts.body).setFontSize(10).setFontStyle('italic').setFontColor('#5f6368')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f8f9fa');
  sheet.setRowHeight(2, 24);
  sheet.setRowHeight(3, 8);

  // KPI band — four compact cards, each spanning two columns
  var monthStart = '(EOMONTH(TODAY(),-1)+1)';
  var monthEnd = 'EOMONTH(TODAY(),0)';
  var cards = [
    { col: 1, label: 'MONTH SPENT', accent: '#188038', fmt: '$#,##0.00',
      value: '=IFERROR(SUMIFS(Expenses!$C$3:$C,Expenses!$H$3:$H,"Expense",Expenses!$A$3:$A,">="&' + monthStart + ',Expenses!$A$3:$A,"<="&' + monthEnd + '),0)' },
    { col: 3, label: 'MONTH EXPENSES', accent: '#1a73e8', fmt: '#,##0',
      value: '=IFERROR(COUNTIFS(Expenses!$H$3:$H,"Expense",Expenses!$A$3:$A,">="&' + monthStart + ',Expenses!$A$3:$A,"<="&' + monthEnd + '),0)' },
    { col: 5, label: 'ALL-TIME SPENT', accent: '#d93025', fmt: '$#,##0.00',
      value: '=IFERROR(SUMIFS(Expenses!$C$3:$C,Expenses!$H$3:$H,"Expense"),0)' },
    { col: 7, label: 'ACTIVE PEOPLE', accent: '#9334e6', fmt: '#,##0',
      value: "=COUNTA('Master Users'!$B$3:$B)" }
  ];
  for (var c = 0; c < cards.length; c++) {
    var cd = cards[c];
    sheet.getRange(4, cd.col, 1, 2).merge().setValue(cd.label)
      .setFontFamily(CONFIG.fonts.header).setFontSize(9).setFontWeight('bold').setFontColor(cd.accent)
      .setBackground('#f8f9fa').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(5, cd.col, 1, 2).merge().setFormula(cd.value)
      .setFontFamily(CONFIG.fonts.header).setFontSize(17).setFontWeight('bold').setFontColor('#202124')
      .setBackground('#f8f9fa').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setNumberFormat(cd.fmt);
    sheet.getRange(4, cd.col, 2, 2).setBorder(true, true, true, true, false, false, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);
  }
  sheet.setRowHeight(4, 20);
  sheet.setRowHeight(5, 40);
  sheet.setRowHeight(6, 10);

  // Helper tip line (full width, above the growing tables)
  sheet.getRange(7, 1, 1, 8).merge()
    .setValue('💡 Net Balance = who is up or down overall.   ✅ Green = owed to them   ❌ Red = they owe.   "Suggested Settlements" lists the fewest payments to settle everyone up.')
    .setFontFamily(CONFIG.fonts.body).setFontSize(10).setFontStyle('italic')
    .setBackground('#e8f0fe').setFontColor('#1967d2').setWrap(true)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(7, 32);
  sheet.setRowHeight(8, 8);

  // Side-by-side section headers
  sheet.getRange(9, 1, 1, 5).merge().setValue('👥 SPENDING BY PERSON')
    .setFontFamily(CONFIG.fonts.header).setFontSize(13).setFontWeight('bold')
    .setBackground('#ea4335').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(9, 6, 1, 3).merge().setValue('💸 SUGGESTED SETTLEMENTS')
    .setFontFamily(CONFIG.fonts.header).setFontSize(13).setFontWeight('bold')
    .setBackground('#1a73e8').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(9, 32);

  // Table headers
  sheet.getRange(DASH.personHeaderRow, 1, 1, 5).setValues([['Person', 'Total Spent', '# Transactions', 'Net Balance', 'Status']])
    .setFontFamily(CONFIG.fonts.header).setFontWeight('bold').setFontSize(10)
    .setBackground('#f1f3f4').setHorizontalAlignment('center');
  sheet.getRange(DASH.sugHeaderRow, DASH.sugFirstCol, 1, 3).setValues([['Who Pays', 'Who Receives', 'Amount']])
    .setFontFamily(CONFIG.fonts.header).setFontWeight('bold').setFontSize(10)
    .setBackground('#f1f3f4').setHorizontalAlignment('center');
  sheet.setRowHeight(DASH.personHeaderRow, 30);

  // Per-person rows (left region)
  var row = DASH.personDataRow;
  for (var i = 0; i < userArray.length; i++) {
    writeDashboardPersonRow(sheet, row, userArray[i], CONFIG.colors.userColors[i % CONFIG.colors.userColors.length]);
    row++;
  }
  if (userArray.length > 0) {
    sheet.getRange(DASH.personHeaderRow, 1, userArray.length + 1, 5)
      .setBorder(true, true, true, true, true, true, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);
  }

  if (sheet.getMaxColumns() > 8) sheet.hideColumns(9, sheet.getMaxColumns() - 8);
  lockSheet(sheet, 'Dashboard');
  sheet.setFrozenRows(2);
}

/** Writes the five formula-driven cells for one Dashboard person row (A–E). */
function writeDashboardPersonRow(sheet, row, name, color) {
  sheet.getRange(row, 1).setValue(name).setFontFamily(CONFIG.fonts.body).setFontSize(10).setFontWeight('bold').setBackground(color);
  sheet.getRange(row, 2)
    .setFormula('=IFERROR(SUMIFS(Ledger!$C$3:$C,Ledger!$B$3:$B,$A' + row + ',Ledger!$E$3:$E,"Expense"),0)')
    .setNumberFormat('$#,##0.00').setFontSize(10).setHorizontalAlignment('center').setBackground(color);
  sheet.getRange(row, 3)
    .setFormula('=IFERROR(COUNTIFS(Expenses!$D$3:$D,$A' + row + ',Expenses!$H$3:$H,"Expense"),0)')
    .setNumberFormat('#,##0').setFontSize(10).setHorizontalAlignment('center').setBackground(color);
  sheet.getRange(row, 4)
    .setFormula('=IFERROR(SUMIFS(Ledger!$C$3:$C,Ledger!$B$3:$B,$A' + row + ')-SUMIFS(Ledger!$D$3:$D,Ledger!$B$3:$B,$A' + row + '),0)')
    .setNumberFormat('$#,##0.00').setFontSize(10).setHorizontalAlignment('center').setFontWeight('bold').setBackground(color);
  sheet.getRange(row, 5)
    .setFormula('=IF($D' + row + '>0.005,"✅ Owed $"&TEXT($D' + row + ',"#,##0.00"),IF($D' + row + '<-0.005,"❌ Owes $"&TEXT(-$D' + row + ',"#,##0.00"),"Even"))')
    .setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold').setBackground(color);
  sheet.setRowHeight(row, 30);
}

/* Balances is merged into the Dashboard — see createDashboardSheet (left region)
 * and writeSettlementSuggestions (right region). There is no separate Balances sheet. */

/* ============================================================
 *  SHARED SMALL HELPERS
 * ============================================================ */

function sectionHeader(sheet, row, span, title, bg) {
  sheet.getRange(row, 1, 1, span).merge().setValue(title)
    .setFontFamily(CONFIG.fonts.header).setFontSize(15).setFontWeight('bold')
    .setBackground(bg).setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.setRowHeight(row, 40);
}

function keyValueRows(sheet, row, rows, bg) {
  for (var i = 0; i < rows.length; i++) {
    sheet.getRange(row, 1).setValue(rows[i][0]).setFontFamily(CONFIG.fonts.body).setFontWeight('bold').setFontSize(12).setBackground(bg);
    sheet.getRange(row, 2).setFormula(rows[i][1]).setFontFamily(CONFIG.fonts.body).setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center').setNumberFormat(rows[i][2]);
    sheet.getRange(row, 1, 1, 2).setBorder(true, true, true, true, false, false, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(row, 38);
    row++;
  }
  return row;
}

function lockSheet(sheet, label) {
  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < existing.length; i++) existing[i].remove();
  var protection = sheet.protect();
  protection.setDescription(label + ' - Protected');
  protection.setWarningOnly(false);
}

function unlockSheet(sheet) {
  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < existing.length; i++) existing[i].remove();
}

function getNamesFromArray(userArray) {
  return userArray.map(function (u) { return (typeof u === 'string') ? u.trim() : u; });
}

function getAllUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheets.users);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  return sheet.getRange(3, 2, lastRow - 2, 1).getValues()
    .map(function (r) { return r[0]; })
    .filter(function (u) { return u !== ''; });
}

/* ============================================================
 *  CONDITIONAL FORMATTING (Expenses sheet)
 * ============================================================ */

function applyExpensesConditionalFormatting(sheet, users) {
  var paidByRange = sheet.getRange('D3:D1000');
  var typeRange = sheet.getRange('H3:H1000');
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Expense')
      .setBackground(CONFIG.colors.positiveGreen).setFontColor('#2e7d32').setRanges([typeRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Settlement')
      .setBackground('#bbdefb').setFontColor('#1565c0').setRanges([typeRange]).build()
  ];
  for (var i = 0; i < users.length; i++) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(users[i])
      .setBackground(CONFIG.colors.userColors[i % CONFIG.colors.userColors.length])
      .setRanges([paidByRange]).build());
  }
  sheet.setConditionalFormatRules(rules);
}

function updateExpensesConditionalFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheets.expenses);
  if (!sheet) return;
  applyExpensesConditionalFormatting(sheet, getAllUsers());
}

/* ============================================================
 *  ADD USER
 * ============================================================ */

function addNewUser(userName) {
  try {
    userName = (userName || '').trim();
    if (!userName) return { success: false, message: '❌ Please enter a name.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usersSheet = ss.getSheetByName(CONFIG.sheets.users);
    if (!usersSheet) return { success: false, message: '❌ Master Users sheet not found.' };

    var existing = getAllUsers();
    if (existing.indexOf(userName) !== -1) return { success: false, message: '❌ User "' + userName + '" already exists!' };

    var index = existing.length; // 0-based position of the new user
    var color = CONFIG.colors.userColors[index % CONFIG.colors.userColors.length];

    // --- Master Users ---
    var wasHidden = usersSheet.isSheetHidden();
    if (wasHidden) usersSheet.showSheet();
    unlockSheet(usersSheet);
    var uRow = usersSheet.getLastRow() + 1;
    usersSheet.getRange(uRow, 1).setValue(index + 1).setFontFamily(CONFIG.fonts.body).setHorizontalAlignment('center');
    usersSheet.getRange(uRow, 2).setValue(userName).setFontFamily(CONFIG.fonts.body).setFontWeight('bold');
    usersSheet.getRange(uRow, 3).setValue(color).setBackground(color).setFontFamily(CONFIG.fonts.mono).setFontSize(9);
    usersSheet.setRowHeight(uRow, 32);
    usersSheet.getRange(3, 1, uRow - 2, 3).setBorder(true, true, true, true, true, true, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);
    lockSheet(usersSheet, 'Master Users');
    if (wasHidden) usersSheet.hideSheet();

    // --- Dashboard person table (A-E). Row derived from the fixed header so the
    //     F-H suggestions block can never throw off the placement. ---
    var dash = ss.getSheetByName(CONFIG.sheets.dashboard);
    if (dash) {
      unlockSheet(dash);
      var dRow = DASH.personHeaderRow + index + 1; // index = count of existing users
      writeDashboardPersonRow(dash, dRow, userName, color);
      dash.getRange(DASH.personHeaderRow, 1, index + 2, 5)
        .setBorder(true, true, true, true, true, true, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);
      lockSheet(dash, 'Dashboard');
    }

    updateExpensesConditionalFormatting();
    rebuildLedger();

    return { success: true, message: '✅ User "' + userName + '" added successfully!' };
  } catch (error) {
    Logger.log('Add user error: ' + error.toString());
    return { success: false, message: '❌ Error: ' + error.toString() };
  }
}

/* ============================================================
 *  ADD EXPENSE / SETTLEMENT  (append to Expenses, then rebuild Ledger)
 * ============================================================ */

function addExpense(expenseData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheets.expenses);
    var newRow = Math.max(sheet.getLastRow(), 2) + 1;

    unlockSheet(sheet);
    sheet.getRange(newRow, 1, 1, 8).setValues([[
      parseLocalDate(expenseData.date),
      expenseData.description,
      parseFloat(expenseData.amount),
      expenseData.paidBy,
      expenseData.splitType,
      expenseData.splitDetails,
      expenseData.category || 'General',
      'Expense'
    ]]);
    sheet.getRange(newRow, 1).setNumberFormat('mmm dd, yyyy');
    sheet.getRange(newRow, 3).setNumberFormat('$#,##0.00');
    lockSheet(sheet, 'Expenses');

    rebuildLedger();
    return { success: true, message: '✅ Expense added successfully!' };
  } catch (error) {
    return { success: false, message: '❌ Error: ' + error.toString() };
  }
}

function addSettlement(settlementData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheets.expenses);
    var newRow = Math.max(sheet.getLastRow(), 2) + 1;

    var details = settlementData.from + ' paid ' + settlementData.to;
    unlockSheet(sheet);
    sheet.getRange(newRow, 1, 1, 8).setValues([[
      new Date(),
      '💳 Settlement: ' + details,
      parseFloat(settlementData.amount),
      settlementData.from,
      'Settlement',
      details,
      'Settlement',
      'Settlement'
    ]]);
    sheet.getRange(newRow, 1).setNumberFormat('mmm dd, yyyy');
    sheet.getRange(newRow, 3).setNumberFormat('$#,##0.00');
    lockSheet(sheet, 'Expenses');

    rebuildLedger();
    return { success: true, message: '✅ Settlement recorded!' };
  } catch (error) {
    return { success: false, message: '❌ Error: ' + error.toString() };
  }
}

/* ============================================================
 *  LEDGER REBUILD  (the heart of the self-healing design)
 * ============================================================ */

/**
 * Reads every row in Expenses and regenerates the Ledger from scratch.
 * Ledger schema: Date | Person | Paid | Share | Type | Description
 *  - Expense: payer gets a Paid row for the full amount; each participant
 *    gets a Share row for their portion.
 *  - Settlement: payer gets Paid=amount; receiver gets Share=amount.
 * Net balance for a person = SUM(Paid) - SUM(Share).
 */
function rebuildLedger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var expenses = ss.getSheetByName(CONFIG.sheets.expenses);
  var ledger = ss.getSheetByName(CONFIG.sheets.ledger);
  if (!expenses || !ledger) return;

  var users = getAllUsers();
  var ledgerRows = [];

  var lastRow = expenses.getLastRow();
  if (lastRow >= 3) {
    var data = expenses.getRange(3, 1, lastRow - 2, 8).getValues();
    for (var i = 0; i < data.length; i++) {
      var date = data[i][0];
      var description = data[i][1];
      var amount = parseFloat(data[i][2]);
      var paidBy = data[i][3];
      var splitType = data[i][4];
      var splitDetails = data[i][5];
      var type = data[i][7];

      if (!amount || !type) continue;

      if (type === 'Settlement') {
        var parts = String(splitDetails).split(' paid ');
        if (parts.length === 2) {
          var from = parts[0].trim();
          var to = parts[1].trim();
          ledgerRows.push([date, from, amount, 0, 'Settlement', description]);
          ledgerRows.push([date, to, 0, amount, 'Settlement', description]);
        }
      } else { // Expense
        ledgerRows.push([date, paidBy, amount, 0, 'Expense', description]);
        var shares = calculateShares(amount, splitType, splitDetails, users);
        // Fallback: a quick-add (e.g. via Canvas inline-add, or typed straight
        // into the grid) may have no/garbled split details. Rather than lose the
        // split, default to an equal split among all current users so balances
        // stay correct. Structured adds via the dialog are unaffected.
        if (Object.keys(shares).length === 0 && users.length > 0) {
          var perHead = amount / users.length;
          for (var u = 0; u < users.length; u++) shares[users[u]] = perHead;
        }
        for (var person in shares) {
          if (shares.hasOwnProperty(person)) {
            ledgerRows.push([date, person, 0, roundMoney(shares[person]), 'Expense', description]);
          }
        }
      }
    }
  }

  // Clear old ledger data and write fresh.
  unlockSheet(ledger);
  var ledgerLast = ledger.getLastRow();
  if (ledgerLast >= 3) ledger.getRange(3, 1, ledgerLast - 2, 6).clearContent();
  if (ledgerRows.length > 0) {
    ledger.getRange(3, 1, ledgerRows.length, 6).setValues(ledgerRows);
    ledger.getRange(3, 1, ledgerRows.length, 1).setNumberFormat('mmm dd, yyyy');
    ledger.getRange(3, 3, ledgerRows.length, 2).setNumberFormat('$#,##0.00');
  }
  lockSheet(ledger, 'Ledger');

  // Derive net balances straight from the ledger rows we just built, then
  // regenerate the "who pays whom" suggestions. Fully recomputed each time,
  // so it can never drift out of sync.
  var netMap = {};
  for (var k = 0; k < ledgerRows.length; k++) {
    var nm = ledgerRows[k][1];
    netMap[nm] = (netMap[nm] || 0) + ledgerRows[k][2] - ledgerRows[k][3];
  }
  writeSettlementSuggestions(ss, computeSettlements(netMap));
}

/**
 * Greedy debt simplification: repeatedly match the biggest debtor with the
 * biggest creditor. Produces at most (n-1) payments and always settles
 * everyone to zero. Returns [{from, to, amount}, ...].
 */
function computeSettlements(netMap) {
  var debtors = [], creditors = [];
  for (var p in netMap) {
    if (!netMap.hasOwnProperty(p)) continue;
    var v = roundMoney(netMap[p]);
    if (v < -0.005) debtors.push({ name: p, amt: -v });
    else if (v > 0.005) creditors.push({ name: p, amt: v });
  }
  debtors.sort(function (a, b) { return b.amt - a.amt; });
  creditors.sort(function (a, b) { return b.amt - a.amt; });

  var res = [], i = 0, j = 0, guard = 0;
  while (i < debtors.length && j < creditors.length && guard < 10000) {
    guard++;
    var pay = roundMoney(Math.min(debtors[i].amt, creditors[j].amt));
    if (pay > 0.005) res.push({ from: debtors[i].name, to: creditors[j].name, amount: pay });
    debtors[i].amt = roundMoney(debtors[i].amt - pay);
    creditors[j].amt = roundMoney(creditors[j].amt - pay);
    if (debtors[i].amt < 0.005) i++;
    if (creditors[j].amt < 0.005) j++;
  }
  return res;
}

/** Regenerates the F-H suggestions block on the Dashboard sheet. */
function writeSettlementSuggestions(ss, settlements) {
  var dash = ss.getSheetByName(CONFIG.sheets.dashboard);
  if (!dash) return;

  unlockSheet(dash);

  // Clear any previous suggestions (generous range; F-H from data row down).
  dash.getRange(DASH.sugDataRow, DASH.sugFirstCol, 1000, 3).clearContent()
    .setBackground(null).setBorder(false, false, false, false, false, false)
    .setFontWeight('normal').setFontColor('#000000');
  // Un-merge any leftover "all settled" banner from a prior run.
  try { dash.getRange(DASH.sugDataRow, DASH.sugFirstCol, 1, 3).breakApart(); } catch (e) {}

  if (!settlements || settlements.length === 0) {
    dash.getRange(DASH.sugDataRow, DASH.sugFirstCol, 1, 3).merge()
      .setValue('🎉 All settled up — no payments needed!')
      .setFontFamily(CONFIG.fonts.body).setFontWeight('bold').setFontSize(10)
      .setBackground(CONFIG.colors.positiveGreen).setFontColor('#1b5e20')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.setRowHeight(DASH.sugDataRow, 32);
    dash.getRange(DASH.sugHeaderRow, DASH.sugFirstCol, 2, 3)
      .setBorder(true, true, true, true, true, true, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);
    lockSheet(dash, 'Dashboard');
    return;
  }

  var rows = settlements.map(function (s) { return [s.from, '→  ' + s.to, s.amount]; });
  dash.getRange(DASH.sugDataRow, DASH.sugFirstCol, rows.length, 3).setValues(rows);
  dash.getRange(DASH.sugDataRow, DASH.sugFirstCol, rows.length, 1)
    .setFontFamily(CONFIG.fonts.body).setFontSize(10).setFontWeight('bold').setBackground(CONFIG.colors.negativeRed).setHorizontalAlignment('center');
  dash.getRange(DASH.sugDataRow, DASH.sugFirstCol + 1, rows.length, 1)
    .setFontFamily(CONFIG.fonts.body).setFontSize(10).setFontWeight('bold').setBackground(CONFIG.colors.positiveGreen).setHorizontalAlignment('center');
  dash.getRange(DASH.sugDataRow, DASH.sugFirstCol + 2, rows.length, 1)
    .setFontFamily(CONFIG.fonts.body).setFontSize(10).setFontWeight('bold').setNumberFormat('$#,##0.00').setHorizontalAlignment('center');
  for (var r = 0; r < rows.length; r++) dash.setRowHeight(DASH.sugDataRow + r, 30);

  dash.getRange(DASH.sugHeaderRow, DASH.sugFirstCol, rows.length + 1, 3)
    .setBorder(true, true, true, true, true, true, CONFIG.colors.borderColor, SpreadsheetApp.BorderStyle.SOLID);

  lockSheet(dash, 'Dashboard');
}

function refreshBalances() {
  try {
    rebuildLedger();
    SpreadsheetApp.flush();
    return { success: true, message: '✅ Balances fully recalculated from your expenses!' };
  } catch (error) {
    return { success: false, message: '❌ Error: ' + error.toString() };
  }
}

function refreshBalancesWithUI() {
  SpreadsheetApp.getUi().alert(refreshBalances().message);
}

/* ============================================================
 *  SPLIT MATH  (server-side, reliable)
 * ============================================================ */

function calculateShares(amount, splitType, splitDetails, users) {
  var shares = {};
  try {
    if (splitType === 'Equal') {
      var m = String(splitDetails).match(/\[(.*?)\]/);
      if (m && m[1]) {
        var list = m[1].split(',').map(function (u) { return u.trim(); }).filter(function (u) { return u !== ''; });
        var per = amount / list.length;
        list.forEach(function (u) { shares[u] = (shares[u] || 0) + per; });
      }
    } else if (splitType === 'Exact') {
      var ex = String(splitDetails).match(/([^:,]+):\s*\$?([\d.]+)/g);
      if (ex) ex.forEach(function (item) {
        var p = item.split(':');
        shares[p[0].trim()] = parseFloat(p[1].replace('$', '').trim());
      });
    } else if (splitType === 'Percentage') {
      var pc = String(splitDetails).match(/([^:,]+):\s*([\d.]+)%/g);
      if (pc) pc.forEach(function (item) {
        var p = item.split(':');
        shares[p[0].trim()] = amount * (parseFloat(p[1].replace('%', '').trim()) / 100);
      });
    } else if (splitType === 'Shares') {
      var sh = String(splitDetails).match(/([^:,]+):\s*([\d.]+)\s*share/gi);
      if (sh) {
        var total = 0, tmp = {};
        sh.forEach(function (item) {
          var p = item.split(':');
          var v = parseFloat(p[1].replace(/share/i, '').trim());
          tmp[p[0].trim()] = v; total += v;
        });
        for (var u in tmp) if (tmp.hasOwnProperty(u)) shares[u] = amount * (tmp[u] / total);
      }
    }
  } catch (e) {
    Logger.log('calculateShares error: ' + e);
  }
  return shares;
}

/* ============================================================
 *  ARCHIVE
 * ============================================================ */

function archiveExpenses(beforeDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var expenses = ss.getSheetByName(CONFIG.sheets.expenses);
    var archives = ss.getSheetByName(CONFIG.sheets.archives);
    if (!archives) { createArchivesSheet(); archives = ss.getSheetByName(CONFIG.sheets.archives); }

    var cutoff = parseLocalDate(beforeDate);
    var today = new Date();
    var lastRow = expenses.getLastRow();
    if (lastRow < 3) return { success: true, message: 'No expenses to archive', count: 0 };

    var data = expenses.getRange(3, 1, lastRow - 2, 8).getValues();
    var toArchive = [], toKeep = [];
    for (var i = 0; i < data.length; i++) {
      if (new Date(data[i][0]) < cutoff) toArchive.push(data[i].concat([today]));
      else toKeep.push(data[i]);
    }

    if (toArchive.length > 0) {
      var aLast = archives.getLastRow();
      unlockSheet(archives);
      archives.getRange(aLast + 1, 1, toArchive.length, 9).setValues(toArchive);
      archives.getRange(aLast + 1, 1, toArchive.length, 1).setNumberFormat('mmm dd, yyyy');
      archives.getRange(aLast + 1, 3, toArchive.length, 1).setNumberFormat('$#,##0.00');
      archives.getRange(aLast + 1, 9, toArchive.length, 1).setNumberFormat('mmm dd, yyyy');
      lockSheet(archives, 'Archives');

      unlockSheet(expenses);
      expenses.getRange(3, 1, lastRow - 2, 8).clearContent();
      if (toKeep.length > 0) {
        expenses.getRange(3, 1, toKeep.length, 8).setValues(toKeep);
        expenses.getRange(3, 1, toKeep.length, 1).setNumberFormat('mmm dd, yyyy');
        expenses.getRange(3, 3, toKeep.length, 1).setNumberFormat('$#,##0.00');
      }
      lockSheet(expenses, 'Expenses');

      rebuildLedger();
    }
    return { success: true, message: '✅ ' + toArchive.length + ' expense(s) archived!', count: toArchive.length };
  } catch (error) {
    return { success: false, message: '❌ Error: ' + error.toString() };
  }
}

/* ============================================================
 *  UTILITIES
 * ============================================================ */

/** Parse a yyyy-mm-dd string into a LOCAL date (avoids UTC off-by-one). */
function parseLocalDate(value) {
  if (value instanceof Date) return value;
  var s = String(value);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  return new Date(value);
}

function roundMoney(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}
