/**
 * TransactionManager.gs
 * ---------------------
 * Core transaction management system for a small business running on Google Sheets.
 *
 * What it does:
 *  - Adds a custom "Transactions" menu to the spreadsheet UI.
 *  - Opens a dialog form for adding a transaction (type, date, category, amount, note).
 *  - Validates every entry before it is written (date format, positive amount, allowed type).
 *  - Auto-categorizes transactions based on keyword rules when the user leaves
 *    the category blank or picks "Auto".
 *  - Generates a monthly summary into a "Dashboard" sheet (delegated to Dashboard.gs).
 *
 * Sheet layout it expects ("Transactions" sheet):
 *   A: Date | B: Type (Income/Expense) | C: Category | D: Amount | E: Note | F: CreatedBy
 *
 * Setup:
 *  1. Open the Google Sheet, then Extensions > Apps Script.
 *  2. Paste this file plus ExpenseTracker.gs and Dashboard.gs into the project.
 *  3. Save, then run setupMenu() once (or just reload the sheet).
 *  4. Use the new "Transactions" menu.
 */

/* ----------------------------- Configuration ----------------------------- */

const TM_CONFIG = {
  sheetName: 'Transactions',
  headers: ['Date', 'Type', 'Category', 'Amount', 'Note', 'CreatedBy'],
  allowedTypes: ['Income', 'Expense'],
  // Keyword -> Category rules used by auto-categorization (checked in order).
  autoRules: [
    { keywords: ['salary', 'payroll', 'invoice paid', 'client payment'], category: 'Revenue', type: 'Income' },
    { keywords: ['rent', 'lease'], category: 'Rent', type: 'Expense' },
    { keywords: ['uber', 'ola', 'fuel', 'petrol', 'diesel', 'taxi', 'metro'], category: 'Transport', type: 'Expense' },
    { keywords: ['swiggy', 'zomato', 'restaurant', 'lunch', 'dinner', 'groceries'], category: 'Food', type: 'Expense' },
    { keywords: ['electricity', 'water bill', 'internet', 'phone', 'utility'], category: 'Utilities', type: 'Expense' },
    { keywords: ['amazon', 'flipkart', 'stationery', 'supplies', 'office'], category: 'Office Supplies', type: 'Expense' },
    { keywords: ['ads', 'marketing', 'promotion', 'sponsor'], category: 'Marketing', type: 'Expense' },
    { keywords: ['refund'], category: 'Refunds', type: 'Income' }
  ],
  defaultExpenseCategory: 'Miscellaneous',
  defaultIncomeCategory: 'Other Income'
};

/* ------------------------------- UI Menu --------------------------------- */

/**
 * onOpen runs automatically every time the spreadsheet is opened.
 * It installs the custom menu.
 */
function onOpen() {
  setupMenu();
}

/**
 * Installs the "Transactions" custom menu.
 * Safe to call manually if the menu does not appear.
 */
function setupMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Transactions')
    .addItem('➕ Add Transaction…', 'showAddTransactionDialog')
    .addItem('📊 Generate Monthly Summary', 'generateMonthlySummary')
    .addItem('🧹 Clean Up Sheet (format headers)', 'formatTransactionSheet')
    .addToUi();
}

/* --------------------------- Add-transaction UI --------------------------- */

/**
 * Shows an HTML dialog with a transaction entry form.
 * Uses google.script.run to submit back to addTransactionFromForm().
 */
function showAddTransactionDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 12px; }
      label { display: block; margin-top: 10px; font-weight: bold; font-size: 13px; }
      input, select { width: 100%; padding: 6px; margin-top: 4px; box-sizing: border-box; }
      button { margin-top: 14px; padding: 8px 16px; background: #1a73e8; color: #fff; border: 0; border-radius: 4px; cursor: pointer; }
      #msg { margin-top: 10px; font-size: 13px; }
      .err { color: #c5221f; } .ok { color: #137333; }
    </style>
    <label>Type
      <select id="type">
        <option>Expense</option>
        <option>Income</option>
      </select>
    </label>
    <label>Date <input type="date" id="date"></label>
    <label>Category (leave blank for auto)
      <input type="text" id="category" placeholder="e.g. Food — or leave blank">
    </label>
    <label>Amount <input type="number" id="amount" step="0.01" min="0.01"></label>
    <label>Note <input type="text" id="note" placeholder="What was this for?"></label>
    <button onclick="submit()">Add transaction</button>
    <div id="msg"></div>
    <script>
      document.getElementById('date').valueAsDate = new Date();
      function submit() {
        const data = {
          type: document.getElementById('type').value,
          date: document.getElementById('date').value,
          category: document.getElementById('category').value,
          amount: document.getElementById('amount').value,
          note: document.getElementById('note').value
        };
        google.script.run
          .withSuccessHandler(() => {
            document.getElementById('msg').className = 'ok';
            document.getElementById('msg').textContent = 'Saved ✓';
          })
          .withFailureHandler((err) => {
            document.getElementById('msg').className = 'err';
            document.getElementById('msg').textContent = 'Error: ' + err.message;
          })
          .addTransactionFromForm(data);
      }
    </script>
  `)
    .setWidth(340)
    .setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Transaction');
}

/**
 * Entry point called from the dialog form. Validates, auto-categorizes,
 * then appends the row. Throws on validation failure (surfaced in the dialog).
 * @param {Object} data {type, date, category, amount, note}
 */
function addTransactionFromForm(data) {
  const validation = validateTransaction(data);
  if (!validation.ok) {
    throw new Error(validation.errors.join(' '));
  }
  let category = (data.category || '').trim();
  if (!category || category.toLowerCase() === 'auto') {
    category = autoCategorize(data.note || '', data.type);
  }
  const row = [
    new Date(data.date),
    data.type,
    category,
    Number(data.amount),
    data.note || '',
    Session.getActiveUser().getEmail() || 'unknown'
  ];
  getTransactionsSheet().appendRow(row);
  return { ok: true, category: category };
}

/* ------------------------------ Validation ------------------------------- */

/**
 * Validates a transaction payload.
 * @param {Object} data {type, date, category, amount, note}
 * @returns {{ok: boolean, errors: string[]}}
 */
function validateTransaction(data) {
  const errors = [];
  if (!data) {
    return { ok: false, errors: ['No transaction data received.'] };
  }
  if (TM_CONFIG.allowedTypes.indexOf(data.type) === -1) {
    errors.push('Type must be "Income" or "Expense".');
  }
  const d = new Date(data.date);
  if (isNaN(d.getTime())) {
    errors.push('Date is not valid (use YYYY-MM-DD).');
  } else if (d > new Date()) {
    errors.push('Date cannot be in the future.');
  }
  const amount = Number(data.amount);
  if (isNaN(amount) || amount <= 0) {
    errors.push('Amount must be a positive number.');
  }
  if ((data.note || '').length > 280) {
    errors.push('Note is too long (max 280 characters).');
  }
  return { ok: errors.length === 0, errors: errors };
}

/* --------------------------- Auto-categorization -------------------------- */

/**
 * Picks a category from the note text using the keyword rules in TM_CONFIG.
 * Falls back to the default category for the transaction type.
 * @param {string} note Free-text note.
 * @param {string} type 'Income' or 'Expense'.
 * @returns {string} Category name.
 */
function autoCategorize(note, type) {
  const text = (note || '').toLowerCase();
  for (const rule of TM_CONFIG.autoRules) {
    if (rule.type !== type) continue;
    for (const kw of rule.keywords) {
      if (text.indexOf(kw) !== -1) return rule.category;
    }
  }
  return type === 'Income' ? TM_CONFIG.defaultIncomeCategory : TM_CONFIG.defaultExpenseCategory;
}

/* ------------------------------- Sheet I/O ------------------------------- */

/**
 * Returns the Transactions sheet, creating it with headers if missing.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getTransactionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TM_CONFIG.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(TM_CONFIG.sheetName);
    sheet.appendRow(TM_CONFIG.headers);
  }
  return sheet;
}

/**
 * Formats the Transactions sheet: bold frozen header row, date/number formats.
 * Safe to re-run.
 */
function formatTransactionSheet() {
  const sheet = getTransactionsSheet();
  const header = sheet.getRange(1, 1, 1, TM_CONFIG.headers.length);
  header.setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd'); // Date
  sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('#,##0.00');   // Amount
  sheet.autoResizeColumns(1, TM_CONFIG.headers.length);
}

/**
 * Reads all transaction rows as objects.
 * @returns {Array<{date: Date, type: string, category: string, amount: number, note: string}>}
 */
function readTransactions() {
  const sheet = getTransactionsSheet();
  const values = sheet.getDataRange().getValues();
  values.shift(); // drop header
  return values
    .filter((r) => r[0] !== '' && r[0] != null)
    .map((r) => ({
      date: new Date(r[0]),
      type: String(r[1]),
      category: String(r[2]),
      amount: Number(r[3]) || 0,
      note: String(r[4] || '')
    }));
}

/* ---------------------------- Monthly summary ----------------------------- */

/**
 * Menu entry: aggregates the current month's transactions and writes the
 * results into the Dashboard sheet (built by Dashboard.gs).
 */
function generateMonthlySummary() {
  const now = new Date();
  const summary = summarizeMonth(now.getFullYear(), now.getMonth());
  // Dashboard.gs owns the rendering; this keeps one source of truth.
  writeMonthlySummaryToDashboard(summary);
  SpreadsheetApp.getUi().alert(
    'Monthly summary for ' + summary.label + ' written to the Dashboard sheet.'
  );
}

/**
 * Aggregates income/expense totals per category for a given month.
 * @param {number} year Full year, e.g. 2026.
 * @param {number} month Zero-based month index.
 * @returns {{label: string, income: number, expense: number, net: number, byCategory: Object}}
 */
function summarizeMonth(year, month) {
  const rows = readTransactions().filter(
    (t) => t.date.getFullYear() === year && t.date.getMonth() === month
  );
  const byCategory = {};
  let income = 0;
  let expense = 0;
  for (const t of rows) {
    if (t.type === 'Income') income += t.amount;
    else if (t.type === 'Expense') expense += t.amount;
    const key = t.type + ' | ' + t.category;
    byCategory[key] = (byCategory[key] || 0) + t.amount;
  }
  const label = Utilities.formatDate(new Date(year, month, 1), Session.getScriptTimeZone(), 'MMM yyyy');
  return { label: label, income: income, expense: expense, net: income - expense, byCategory: byCategory };
}
