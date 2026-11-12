/**
 * Smart Expense Tracker — Google Sheets + Apps Script
 * Built by Pablo Chomcobar
 *
 * 100% free: runs inside your own Google account.
 * No paid APIs, no servers, no subscriptions.
 *
 * Features
 *  - One-click setup: Expenses, Settings and Dashboard sheets
 *  - Sidebar to add expenses inside the sheet
 *  - Google Form for adding expenses from your phone
 *  - Live dashboard with charts (today, this month, budget, categories, 6 months)
 *  - Budget alert emails (at your % limit and at 100%)
 *  - Automatic monthly report email on the 1st of every month
 */

const CONFIG = {
  SHEETS: {
    EXPENSES: 'Expenses',
    SETTINGS: 'Settings',
    DASHBOARD: 'Dashboard'
  },
  HEADERS: ['Date', 'Category', 'Amount (₹)', 'Payment Mode', 'Note', 'Added On'],
  DEFAULT_CATEGORIES: ['Food', 'Travel', 'Petrol', 'Rent', 'Bills', 'Shopping',
    'Health', 'Education', 'Entertainment', 'Other'],
  DEFAULT_MODES: ['UPI', 'Cash', 'Card', 'Net Banking'],
  DEFAULT_BUDGET: 10000,
  DEFAULT_ALERT_PCT: 80,
  CURRENCY_FORMAT: '₹#,##,##0.00',
  DATE_FORMAT: 'dd-mmm-yyyy',
  COLORS: {
    header: '#111827',
    headerText: '#ffffff',
    accent: '#10b981',
    soft: '#ecfdf5',
    warn: '#f59e0b',
    danger: '#ef4444'
  },
  BRAND: 'Pablo Chomcobar'
};

/* ------------------------------------------------------------------ */
/* Menu                                                               */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Expense Tracker')
    .addItem('➕ Add expense', 'showAddExpenseSidebar')
    .addSeparator()
    .addItem('⚙️ Set up tracker (run once)', 'setupTracker')
    .addItem('📱 Create mobile entry form', 'createEntryForm')
    .addItem('⏰ Turn on automations', 'installAutomations')
    .addItem('⛔ Turn off automations', 'removeAutomations')
    .addSeparator()
    .addItem('📧 Email me this month\'s report now', 'sendCurrentMonthReport')
    .addToUi();
}

/* ------------------------------------------------------------------ */
/* Setup                                                              */
/* ------------------------------------------------------------------ */

function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Kolkata');

  const settings = setupSettingsSheet_(ss);
  setupExpensesSheet_(ss, settings);
  const dashboard = setupDashboardSheet_(ss);

  // Remove the empty default sheet if it is still there
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(blank);
  }

  ss.setActiveSheet(dashboard);
  SpreadsheetApp.getUi().alert(
    '✅ Your tracker is ready!\n\n' +
    '1. Set your monthly budget and email in the "Settings" sheet.\n' +
    '2. Add expenses from the menu: 💰 Expense Tracker → ➕ Add expense.\n' +
    '3. Turn on email alerts: 💰 Expense Tracker → ⏰ Turn on automations.'
  );
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader_(range) {
  range
    .setBackground(CONFIG.COLORS.header)
    .setFontColor(CONFIG.COLORS.headerText)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
}

function setupSettingsSheet_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.SETTINGS);

  // Only fill defaults the first time, never overwrite user changes
  if (!sh.getRange('A1').getValue()) {
    sh.getRange('A1:B1').setValues([['Categories', 'Payment Modes']]);
    sh.getRange(2, 1, CONFIG.DEFAULT_CATEGORIES.length, 1)
      .setValues(CONFIG.DEFAULT_CATEGORIES.map(c => [c]));
    sh.getRange(2, 2, CONFIG.DEFAULT_MODES.length, 1)
      .setValues(CONFIG.DEFAULT_MODES.map(m => [m]));

    sh.getRange('D1:E1').setValues([['Setting', 'Value']]);
    sh.getRange('D2:E4').setValues([
      ['Monthly Budget (₹)', CONFIG.DEFAULT_BUDGET],
      ['Report Email', Session.getEffectiveUser().getEmail() || ''],
      ['Budget alert at (%)', CONFIG.DEFAULT_ALERT_PCT]
    ]);
    sh.getRange('D6').setValue('Tip: add or remove categories / payment modes in columns A and B. ' +
      'The dropdowns update automatically.');
  }

  styleHeader_(sh.getRange('A1:B1'));
  styleHeader_(sh.getRange('D1:E1'));
  sh.getRange('D2:D4').setFontWeight('bold');
  sh.getRange('E2').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('D6').setFontColor('#6b7280').setFontStyle('italic');
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 30);
  sh.setColumnWidth(4, 180);
  sh.setColumnWidth(5, 240);
  sh.setFrozenRows(1);
  return sh;
}

function setupExpensesSheet_(ss, settings) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.EXPENSES);

  if (!sh.getRange('A1').getValue()) {
    sh.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  }
  styleHeader_(sh.getRange(1, 1, 1, CONFIG.HEADERS.length));
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);

  const rows = Math.max(sh.getMaxRows() - 1, 999);
  if (sh.getMaxRows() < rows + 1) sh.insertRowsAfter(sh.getMaxRows(), rows + 1 - sh.getMaxRows());

  sh.getRange(2, 1, rows, 1).setNumberFormat(CONFIG.DATE_FORMAT);
  sh.getRange(2, 3, rows, 1).setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange(2, 6, rows, 1).setNumberFormat('dd-mmm-yyyy hh:mm');

  // Dropdowns pulled live from the Settings sheet
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settings.getRange('A2:A100'), true)
    .setAllowInvalid(false)
    .build();
  const modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settings.getRange('B2:B50'), true)
    .setAllowInvalid(false)
    .build();
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Enter a valid date')
    .build();
  const amountRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0)
    .setAllowInvalid(false)
    .setHelpText('Amount must be more than 0')
    .build();

  sh.getRange(2, 1, rows, 1).setDataValidation(dateRule);
  sh.getRange(2, 2, rows, 1).setDataValidation(categoryRule);
  sh.getRange(2, 3, rows, 1).setDataValidation(amountRule);
  sh.getRange(2, 4, rows, 1).setDataValidation(modeRule);

  // Light zebra stripes for readability
  sh.getBandings().forEach(b => b.remove());
  sh.getRange(1, 1, rows + 1, CONFIG.HEADERS.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, true, false);
  styleHeader_(sh.getRange(1, 1, 1, CONFIG.HEADERS.length));

  [110, 140, 120, 130, 260, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  return sh;
}

function setupDashboardSheet_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.DASHBOARD);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);

  sh.clear();
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.clearConditionalFormatRules();

  const E = CONFIG.SHEETS.EXPENSES;
  const S = CONFIG.SHEETS.SETTINGS;
  const monthStart = 'EOMONTH(TODAY(),-1)+1';
  const monthEnd = 'EOMONTH(TODAY(),0)';

  sh.getRange('A1').setValue('💰 Expense Dashboard')
    .setFontSize(20).setFontWeight('bold');
  sh.getRange('A2').setValue('Smart Expense Tracker · built by ' + CONFIG.BRAND)
    .setFontColor('#6b7280').setFontStyle('italic');

  // KPI cards
  const kpis = [
    ['Spent Today', `=SUMIFS(${E}!C:C,${E}!A:A,TODAY())`],
    ['Spent This Month', `=SUMIFS(${E}!C:C,${E}!A:A,">="&${monthStart},${E}!A:A,"<="&${monthEnd})`],
    ['Monthly Budget', `=${S}!E2`],
    ['Budget Left', '=B6-B5'],
    ['Budget Used', '=IFERROR(B5/B6,0)'],
    ['Entries This Month', `=COUNTIFS(${E}!A:A,">="&${monthStart},${E}!A:A,"<="&${monthEnd})`]
  ];
  sh.getRange(4, 1, kpis.length, 2).setValues(kpis);
  sh.getRange('A4:A9').setFontWeight('bold').setBackground(CONFIG.COLORS.soft);
  sh.getRange('B4:B7').setNumberFormat(CONFIG.CURRENCY_FORMAT);
  sh.getRange('B8').setNumberFormat('0%');
  sh.getRange('B9').setNumberFormat('0');
  sh.getRange('B4:B9').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('A4:B9').setBorder(true, true, true, true, true, true, '#d1d5db',
    SpreadsheetApp.BorderStyle.SOLID);

  // Category table (this month)
  sh.getRange('A12:B12').setValues([['Category', 'This Month']]);
  styleHeader_(sh.getRange('A12:B12'));
  sh.getRange('A13').setFormula(`=FILTER(${S}!A2:A100,${S}!A2:A100<>"")`);
  sh.getRange('B13').setFormula(
    `=ARRAYFORMULA(IF(A13:A60="","",SUMIFS(${E}!C:C,${E}!B:B,A13:A60,` +
    `${E}!A:A,">="&${monthStart},${E}!A:A,"<="&${monthEnd})))`);
  sh.getRange('B13:B60').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  // Last 6 months
  sh.getRange('D12:E12').setValues([['Month', 'Total Spent']]);
  styleHeader_(sh.getRange('D12:E12'));
  sh.getRange('D13').setFormula('=ARRAYFORMULA(EOMONTH(TODAY(),SEQUENCE(6,1,-6,1))+1)');
  sh.getRange('E13').setFormula(
    `=ARRAYFORMULA(SUMIFS(${E}!C:C,${E}!A:A,">="&D13:D18,${E}!A:A,"<="&EOMONTH(D13:D18,0)))`);
  sh.getRange('D13:D18').setNumberFormat('mmm yyyy');
  sh.getRange('E13:E18').setNumberFormat(CONFIG.CURRENCY_FORMAT);

  // Colour the "Budget Used" cell: orange from alert %, red at 100%
  const used = sh.getRange('B8');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(1)
      .setBackground(CONFIG.COLORS.danger).setFontColor('#ffffff')
      .setRanges([used]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=B8>=${S}!E4/100`)
      .setBackground(CONFIG.COLORS.warn).setFontColor('#ffffff')
      .setRanges([used]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor(CONFIG.COLORS.danger)
      .setRanges([sh.getRange('B7')]).build()
  ];
  sh.setConditionalFormatRules(rules);

  [190, 150, 30, 120, 140].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setHiddenGridlines(true);

  // Charts
  SpreadsheetApp.flush();
  const pie = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('A12:B40'))
    .setNumHeaders(1)
    .setPosition(1, 7, 0, 0)
    .setOption('title', 'Where your money went this month')
    .setOption('pieHole', 0.45)
    .setOption('width', 480)
    .setOption('height', 300)
    .build();
  const bars = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('D12:E18'))
    .setNumHeaders(1)
    .setPosition(17, 7, 0, 0)
    .setOption('title', 'Last 6 months')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [CONFIG.COLORS.accent])
    .setOption('width', 480)
    .setOption('height', 300)
    .build();
  sh.insertChart(pie);
  sh.insertChart(bars);
  return sh;
}

/* ------------------------------------------------------------------ */
/* Sidebar entry                                                      */
/* ------------------------------------------------------------------ */

function showAddExpenseSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('AddExpense')
    .setTitle('Add Expense');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Called by the sidebar to fill its dropdowns. */
function getFormOptions() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sh) throw new Error('Run "Set up tracker" first.');
  const last = Math.max(sh.getLastRow(), 2);
  const cols = sh.getRange(2, 1, last - 1, 2).getValues();
  const categories = cols.map(r => String(r[0]).trim()).filter(Boolean);
  const modes = cols.map(r => String(r[1]).trim()).filter(Boolean);
  return { categories: categories, modes: modes, today: todayIso_() };
}

/** Called by the sidebar when the user presses "Save". */
function addExpense(data) {
  const amount = Number(data.amount);
  if (!data.date) throw new Error('Please pick a date.');
  if (!data.category) throw new Error('Please pick a category.');
  if (!(amount > 0)) throw new Error('Amount must be more than 0.');

  const date = parseIsoDate_(data.date);
  appendExpense_(date, data.category, amount, data.mode || '', data.note || '');
  checkBudget_();

  const now = new Date();
  const total = monthTotal_(now.getFullYear(), now.getMonth());
  return 'Saved ' + formatINR_(amount) + ' · This month: ' + formatINR_(total);
}

function appendExpense_(date, category, amount, mode, note) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EXPENSES);
  if (!sh) throw new Error('Run "Set up tracker" first.');
  // Write to the first empty row (appendRow would skip past the formatted area)
  const lastRow = getLastDataRow_(sh);
  sh.getRange(lastRow + 1, 1, 1, 6)
    .setValues([[date, category, amount, mode, String(note).trim(), new Date()]]);
}

function getLastDataRow_(sh) {
  const colA = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] !== '' && colA[i][0] !== null) return i + 1;
  }
  return 1;
}

/* ------------------------------------------------------------------ */
/* Mobile form                                                        */
/* ------------------------------------------------------------------ */

function createEntryForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const existingId = props.getProperty('FORM_ID');
  if (existingId) {
    try {
      const existing = FormApp.openById(existingId);
      ensureFormTrigger_(ss);
      showLinkDialog_('Your mobile form already exists', existing.getPublishedUrl());
      return;
    } catch (e) {
      props.deleteProperty('FORM_ID'); // form was deleted, make a new one
    }
  }

  const opts = getFormOptions();
  const form = FormApp.create('Add Expense — ' + ss.getName());
  form.setDescription('Quick expense entry from your phone. Smart Expense Tracker by ' + CONFIG.BRAND + '.');
  form.setConfirmationMessage('✅ Expense saved!');
  form.setShowLinkToRespondAgain(true);

  form.addDateItem().setTitle('Date').setRequired(true);
  form.addListItem().setTitle('Category').setChoiceValues(opts.categories).setRequired(true);
  form.addTextItem().setTitle('Amount').setRequired(true)
    .setValidation(FormApp.createTextValidation()
      .setHelpText('Enter an amount more than 0')
      .requireNumberGreaterThan(0)
      .build());
  form.addMultipleChoiceItem().setTitle('Payment Mode').setChoiceValues(opts.modes).setRequired(true);
  form.addParagraphTextItem().setTitle('Note');

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  props.setProperty('FORM_ID', form.getId());
  ensureFormTrigger_(ss);

  showLinkDialog_('📱 Mobile form created!', form.getPublishedUrl());
  ui.alert('Tip: open the link on your phone and "Add to Home screen" for one-tap entry.\n\n' +
    'Note: if you add new categories later, run "Create mobile entry form" again ' +
    'after deleting the old form, or edit the form choices by hand.');
}

function ensureFormTrigger_(ss) {
  const exists = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'onFormSubmitHandler');
  if (!exists) {
    ScriptApp.newTrigger('onFormSubmitHandler').forSpreadsheet(ss).onFormSubmit().create();
  }
}

/** Runs automatically whenever the mobile form is submitted. */
function onFormSubmitHandler(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = e.range.getValues()[0];
  const get = name => row[headers.indexOf(name)];

  const amount = Number(get('Amount'));
  if (!(amount > 0)) return;

  let date = get('Date');
  if (!(date instanceof Date)) date = new Date();

  appendExpense_(date, get('Category'), amount, get('Payment Mode') || '', get('Note') || '');
  checkBudget_();
}

function showLinkDialog_(title, url) {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:8px">' +
    '<p>Open this link on your phone:</p>' +
    '<p><a href="' + url + '" target="_blank">' + url + '</a></p>' +
    '</div>').setWidth(460).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

/* ------------------------------------------------------------------ */
/* Automations (time triggers)                                        */
/* ------------------------------------------------------------------ */

const TIME_JOBS_ = ['monthlyReportJob', 'dailyBudgetCheck'];

function installAutomations() {
  removeTimeTriggers_();
  ScriptApp.newTrigger('monthlyReportJob').timeBased().onMonthDay(1).atHour(9).create();
  ScriptApp.newTrigger('dailyBudgetCheck').timeBased().everyDays(1).atHour(21).create();
  SpreadsheetApp.getUi().alert(
    '⏰ Automations are ON\n\n' +
    '• Every day at ~9 PM: budget check (email only if you cross your limit)\n' +
    '• 1st of every month at ~9 AM: last month\'s report by email\n\n' +
    'Emails go to: ' + getReportEmail_());
}

function removeAutomations() {
  const n = removeTimeTriggers_();
  SpreadsheetApp.getUi().alert('⛔ Automations turned off (' + n + ' removed).');
}

function removeTimeTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (TIME_JOBS_.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  return n;
}

function dailyBudgetCheck() {
  checkBudget_();
}

function monthlyReportJob() {
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  sendReport_(prev.getFullYear(), prev.getMonth());
}

function sendCurrentMonthReport() {
  const now = new Date();
  sendReport_(now.getFullYear(), now.getMonth());
  SpreadsheetApp.getUi().alert('📧 Report sent to ' + getReportEmail_());
}

/* ------------------------------------------------------------------ */
/* Budget alerts                                                      */
/* ------------------------------------------------------------------ */

function checkBudget_() {
  const settings = getSettings_();
  if (!(settings.budget > 0) || !settings.email) return;

  const now = new Date();
  const total = monthTotal_(now.getFullYear(), now.getMonth());
  const used = total / settings.budget;
  const key = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM');
  const props = PropertiesService.getDocumentProperties();

  const levels = [
    { id: '100', limit: 1, title: '🚨 Budget crossed!' },
    { id: 'warn', limit: settings.alertPct / 100, title: '⚠️ Budget warning' }
  ];

  for (let i = 0; i < levels.length; i++) {
    const lv = levels[i];
    const flag = 'ALERT_' + key + '_' + lv.id;
    if (used >= lv.limit) {
      if (!props.getProperty(flag)) {
        props.setProperty(flag, '1');
        if (lv.id === '100') props.setProperty('ALERT_' + key + '_warn', '1');
        MailApp.sendEmail({
          to: settings.email,
          subject: lv.title + ' — ' + formatINR_(total) + ' of ' + formatINR_(settings.budget),
          htmlBody: emailShell_(lv.title,
            '<p>You have spent <b>' + formatINR_(total) + '</b> this month, which is <b>' +
            Math.round(used * 100) + '%</b> of your ' + formatINR_(settings.budget) + ' budget.</p>' +
            '<p>Budget left: <b>' + formatINR_(settings.budget - total) + '</b></p>')
        });
      }
      return; // only send the highest alert reached
    }
  }
}

/* ------------------------------------------------------------------ */
/* Reports                                                            */
/* ------------------------------------------------------------------ */

function sendReport_(year, month) {
  const settings = getSettings_();
  if (!settings.email) throw new Error('Add your email in Settings (cell E3).');

  const rows = getExpenseRows_().filter(r =>
    r.date.getFullYear() === year && r.date.getMonth() === month);
  const monthName = Utilities.formatDate(new Date(year, month, 1), 'Asia/Kolkata', 'MMMM yyyy');

  const total = rows.reduce((s, r) => s + r.amount, 0);
  const byCat = groupSum_(rows, 'category');
  const byMode = groupSum_(rows, 'mode');
  const biggest = rows.slice().sort((a, b) => b.amount - a.amount)[0];

  const table = (title, map) => {
    const entries = Object.keys(map).sort((a, b) => map[b] - map[a]);
    if (!entries.length) return '';
    return '<h3 style="margin:18px 0 6px">' + title + '</h3>' +
      '<table style="border-collapse:collapse;width:100%">' +
      entries.map(k =>
        '<tr><td style="padding:6px;border-bottom:1px solid #eee">' + escapeHtml_(k || '—') + '</td>' +
        '<td style="padding:6px;border-bottom:1px solid #eee;text-align:right">' + formatINR_(map[k]) + '</td>' +
        '<td style="padding:6px;border-bottom:1px solid #eee;text-align:right;color:#6b7280">' +
        (total ? Math.round(map[k] / total * 100) : 0) + '%</td></tr>').join('') +
      '</table>';
  };

  let body;
  if (!rows.length) {
    body = '<p>No expenses were recorded in ' + monthName + '.</p>';
  } else {
    const budgetLine = settings.budget > 0
      ? '<p>Budget: ' + formatINR_(settings.budget) + ' · ' +
        (total <= settings.budget
          ? '<span style="color:#10b981">Saved ' + formatINR_(settings.budget - total) + ' 🎉</span>'
          : '<span style="color:#ef4444">Over by ' + formatINR_(total - settings.budget) + '</span>') + '</p>'
      : '';
    body =
      '<p style="font-size:28px;margin:4px 0"><b>' + formatINR_(total) + '</b></p>' +
      '<p style="color:#6b7280;margin:0">' + rows.length + ' expenses in ' + monthName + '</p>' +
      budgetLine +
      table('By category', byCat) +
      table('By payment mode', byMode) +
      (biggest ? '<p style="margin-top:18px">Biggest expense: <b>' + formatINR_(biggest.amount) +
        '</b> on ' + escapeHtml_(biggest.category) +
        (biggest.note ? ' (' + escapeHtml_(biggest.note) + ')' : '') + '</p>' : '');
  }

  MailApp.sendEmail({
    to: settings.email,
    subject: '📊 Expense report — ' + monthName + ' — ' + formatINR_(total),
    htmlBody: emailShell_('Expense report · ' + monthName, body)
  });
}

function emailShell_(title, inner) {
  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  return '<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;color:#111827">' +
    '<div style="background:#111827;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0">' +
    '<b style="font-size:18px">' + title + '</b></div>' +
    '<div style="border:1px solid #e5e7eb;border-top:0;padding:16px 20px;border-radius:0 0 10px 10px">' +
    inner +
    '<p style="margin-top:22px"><a href="' + url + '" style="background:#10b981;color:#fff;' +
    'padding:10px 16px;border-radius:6px;text-decoration:none">Open tracker</a></p>' +
    '<p style="color:#9ca3af;font-size:12px;margin-top:22px">Smart Expense Tracker · by ' +
    CONFIG.BRAND + '</p></div></div>';
}

/* ------------------------------------------------------------------ */
/* Helpers                                                            */
/* ------------------------------------------------------------------ */

function getSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sh) return { budget: 0, email: '', alertPct: CONFIG.DEFAULT_ALERT_PCT };
  const v = sh.getRange('E2:E4').getValues();
  return {
    budget: Number(v[0][0]) || 0,
    email: String(v[1][0] || '').trim() || Session.getEffectiveUser().getEmail(),
    alertPct: Number(v[2][0]) || CONFIG.DEFAULT_ALERT_PCT
  };
}

function getReportEmail_() {
  return getSettings_().email;
}

function getExpenseRows_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EXPENSES);
  if (!sh) return [];
  const last = getLastDataRow_(sh);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues()
    .filter(r => r[0] instanceof Date && Number(r[2]) > 0)
    .map(r => ({
      date: r[0],
      category: String(r[1]),
      amount: Number(r[2]),
      mode: String(r[3]),
      note: String(r[4])
    }));
}

function monthTotal_(year, month) {
  return getExpenseRows_()
    .filter(r => r.date.getFullYear() === year && r.date.getMonth() === month)
    .reduce((s, r) => s + r.amount, 0);
}

function groupSum_(rows, key) {
  return rows.reduce((acc, r) => {
    acc[r[key]] = (acc[r[key]] || 0) + r.amount;
    return acc;
  }, {});
}

/** Formats a number the Indian way: ₹1,23,456.00 */
function formatINR_(n) {
  const neg = n < 0;
  const parts = Math.abs(Number(n) || 0).toFixed(2).split('.');
  let int = parts[0];
  if (int.length > 3) {
    const last3 = int.slice(-3);
    const rest = int.slice(0, -3).replace(/\B(?=(\d{2})+(?!\d))/g, ',');
    int = rest + ',' + last3;
  }
  return (neg ? '-₹' : '₹') + int + '.' + parts[1];
}

function todayIso_() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
}

function parseIsoDate_(iso) {
  const p = String(iso).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
