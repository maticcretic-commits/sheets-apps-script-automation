/**
 * Dashboard.gs
 * ------------
 * Builds and refreshes the "Dashboard" sheet: headline totals, a monthly
 * summary table, embedded charts, conditional formatting, and alerts.
 *
 * What it does:
 *  - Rebuilds the Dashboard sheet from scratch (safe to re-run any time).
 *  - Writes headline KPIs: total income, total expenses, net balance.
 *  - Writes the current month's per-category summary (from TransactionManager).
 *  - Adds an EmbeddedChart: pie chart of expenses by category + column chart
 *    of income vs expenses.
 *  - Applies conditional formatting: negative net in red, high-spend
 *    categories highlighted.
 *  - Raises alerts: rows for overdue/high-spend categories above a threshold.
 *
 * Depends on TransactionManager.gs (summarizeMonth, readTransactions) and
 * ExpenseTracker.gs (getCategoryTotals). All three files live in one
 * Apps Script project bound to the same spreadsheet.
 */

const DB_CONFIG = {
  sheetName: 'Dashboard',
  // Flag a category when this month's spend exceeds it (in sheet currency units).
  highSpendThreshold: 50000,
  alertColor: '#fce8e6',
  okColor: '#e6f4ea'
};

/* ------------------------------ Entry point ------------------------------- */

/**
 * Rebuilds the entire Dashboard sheet. Idempotent — clears and rewrites.
 * Can be run from the menu (Transactions > Generate Monthly Summary) or directly.
 */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DB_CONFIG.sheetName);
  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(DB_CONFIG.sheetName);
  }

  let row = 1;
  row = writeKpis(sheet, row);
  row = writeMonthlySummaryTable(sheet, row);
  row = writeExpenseCategoryTable(sheet, row);
  row = writeAlerts(sheet, row);
  addCharts(sheet);
  applyConditionalFormatting(sheet);
  sheet.autoResizeColumns(1, 4);
  SpreadsheetApp.flush();
}

/**
 * Called by TransactionManager.generateMonthlySummary().
 * Keeps the monthly-summary write path in one place.
 * @param {{label: string, income: number, expense: number, net: number, byCategory: Object}} summary
 */
function writeMonthlySummaryToDashboard(summary) {
  refreshDashboard(); // full rebuild keeps every section consistent
}

/* -------------------------------- Sections -------------------------------- */

/**
 * Writes the KPI block: total income, total expenses, net balance.
 * @returns {number} next free row.
 */
function writeKpis(sheet, startRow) {
  const rows = readTransactions();
  let income = 0;
  let expense = 0;
  for (const t of rows) {
    if (t.type === 'Income') income += t.amount;
    else if (t.type === 'Expense') expense += t.amount;
  }
  const net = income - expense;

  sheet.getRange(startRow, 1).setValue('📈 DASHBOARD — last updated ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'))
    .setFontWeight('bold').setFontSize(14);

  const kpis = [
    ['Total Income', income],
    ['Total Expenses', expense],
    ['Net Balance', net]
  ];
  sheet.getRange(startRow + 1, 1, 1, 2).setValues([['Metric', 'Amount']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sheet.getRange(startRow + 2, 1, kpis.length, 2).setValues(kpis);
  sheet.getRange(startRow + 2, 2, kpis.length, 1).setNumberFormat('#,##0.00');
  // Tag the Net Balance row so conditional formatting can find it.
  sheet.getRange(startRow + 4, 1).setNote('NET_ROW');
  return startRow + 6;
}

/**
 * Writes the current month's per-category summary table.
 * @returns {number} next free row.
 */
function writeMonthlySummaryTable(sheet, startRow) {
  const now = new Date();
  const summary = summarizeMonth(now.getFullYear(), now.getMonth());
  sheet.getRange(startRow, 1).setValue('🗓 Monthly Summary — ' + summary.label)
    .setFontWeight('bold').setFontSize(12);

  const entries = Object.keys(summary.byCategory).sort();
  const table = [['Type | Category', 'Amount']];
  for (const key of entries) {
    table.push([key, summary.byCategory[key]]);
  }
  if (entries.length === 0) table.push(['(no transactions this month)', 0]);

  sheet.getRange(startRow + 1, 1, table.length, 2).setValues(table);
  sheet.getRange(startRow + 1, 1, 1, 2).setFontWeight('bold')
    .setBackground('#e8f0fe');
  sheet.getRange(startRow + 2, 2, Math.max(table.length - 1, 1), 1)
    .setNumberFormat('#,##0.00');
  return startRow + table.length + 2;
}

/**
 * Writes the expense category totals table (base currency, from ExpenseTracker).
 * @returns {number} next free row.
 */
function writeExpenseCategoryTable(sheet, startRow) {
  sheet.getRange(startRow, 1).setValue('💱 Expense Categories (converted to base)')
    .setFontWeight('bold').setFontSize(12);
  const totals = getCategoryTotals();
  const table = [['Category', 'Total', 'Items']];
  for (const t of totals) table.push([t.category, t.total, t.count]);
  if (totals.length === 0) table.push(['(no expenses logged)', 0, 0]);

  sheet.getRange(startRow + 1, 1, table.length, 3).setValues(table);
  sheet.getRange(startRow + 1, 1, 1, 3).setFontWeight('bold').setBackground('#e6f4ea');
  sheet.getRange(startRow + 2, 2, Math.max(table.length - 1, 1), 1)
    .setNumberFormat('#,##0.00');
  return startRow + table.length + 2;
}

/**
 * Writes the alerts block: high-spend categories and negative net balance.
 * @returns {number} next free row.
 */
function writeAlerts(sheet, startRow) {
  sheet.getRange(startRow, 1).setValue('⚠️ Alerts').setFontWeight('bold').setFontSize(12);
  const alerts = buildAlerts();
  const table = [['Alert']];
  for (const a of alerts) table.push([a]);
  if (alerts.length === 0) table.push(['All clear — no alerts.']);

  const range = sheet.getRange(startRow + 1, 1, table.length, 1);
  range.setValues(table);
  range.setFontWeight('bold').setBackground('#fef7e0');
  if (alerts.length > 0) {
    sheet.getRange(startRow + 2, 1, alerts.length, 1).setBackground(DB_CONFIG.alertColor);
  }
  return startRow + table.length + 2;
}

/**
 * Computes alert strings: categories over the high-spend threshold and a
 * negative overall net balance.
 * @returns {string[]}
 */
function buildAlerts() {
  const alerts = [];
  const now = new Date();
  const summary = summarizeMonth(now.getFullYear(), now.getMonth());
  for (const key of Object.keys(summary.byCategory)) {
    if (key.indexOf('Expense |') === 0 && summary.byCategory[key] > DB_CONFIG.highSpendThreshold) {
      alerts.push('High spend: ' + key.replace('Expense | ', '') +
        ' at ' + summary.byCategory[key].toFixed(2) +
        ' this month (threshold ' + DB_CONFIG.highSpendThreshold + ').');
    }
  }
  if (summary.net < 0) {
    alerts.push('Net balance for ' + summary.label + ' is negative (' +
      summary.net.toFixed(2) + '). Review expenses.');
  }
  return alerts;
}

/* ----------------------------- Charts & rules ----------------------------- */

/**
 * Adds a pie chart of expenses by category and a column chart of
 * income vs expenses for the current month.
 */
function addCharts(sheet) {
  const now = new Date();
  const summary = summarizeMonth(now.getFullYear(), now.getMonth());

  // Find the monthly summary table bounds: locate its header row.
  const values = sheet.getDataRange().getValues();
  let headerRow = -1;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === 'Type | Category') { headerRow = i + 1; break; }
  }
  if (headerRow !== -1) {
    const nRows = Object.keys(summary.byCategory).length || 1;
    const pie = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(headerRow + 1, 1, nRows, 2))
      .setPosition(2, 4, 0, 0)
      .setOption('title', 'This month by category')
      .setOption('width', 420)
      .setOption('height', 300)
      .build();
    sheet.insertChart(pie);
  }

  // Income vs expenses column chart from the KPI block (rows 2-4, cols A:B).
  const bar = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(3, 1, 2, 2))
    .setPosition(20, 4, 0, 0)
    .setOption('title', 'Income vs Expenses (all time)')
    .setOption('width', 420)
    .setOption('height', 300)
    .build();
  sheet.insertChart(bar);
}

/**
 * Applies conditional formatting:
 *  - Net Balance cell red background when negative, green when positive.
 *  - Alert rows already tinted; also flag any amount cell over the threshold.
 */
function applyConditionalFormatting(sheet) {
  const rules = [];
  // Net Balance is the 3rd KPI row: row 4 of the sheet, column B.
  const netCell = sheet.getRange(4, 2);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#fce8e6')
      .setFontColor('#a50e0e')
      .setRanges([netCell])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#e6f4ea')
      .setFontColor('#137333')
      .setRanges([netCell])
      .build()
  );
  // Highlight any monthly amount above the high-spend threshold.
  const lastRow = Math.max(sheet.getLastRow(), 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(DB_CONFIG.highSpendThreshold)
      .setBackground('#fef7e0')
      .setRanges([sheet.getRange(1, 2, lastRow, 1)])
      .build()
  );
  sheet.setConditionalFormatRules(rules);
}
