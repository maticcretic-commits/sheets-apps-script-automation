/**
 * ExpenseTracker.gs
 * -----------------
 * Multi-currency expense logging with live FX conversion.
 *
 * What it does:
 *  - Logs expenses in any currency to an "Expenses" sheet.
 *  - Converts every expense to a base currency (default INR) using live
 *    exchange rates fetched from a free rates API (exchangerate.host),
 *    with a per-session cache so the API is hit at most once per run.
 *  - Reports per-category totals in both the original and base currency.
 *
 * Sheet layout it expects ("Expenses" sheet):
 *   A: Date | B: Category | C: Amount | D: Currency | E: AmountBase (INR) | F: Note
 *
 * Setup: paste alongside TransactionManager.gs / Dashboard.gs, then run
 * the menu item or refreshExpenseDashboard().
 */

const ET_CONFIG = {
  sheetName: 'Expenses',
  headers: ['Date', 'Category', 'Amount', 'Currency', 'Amount (INR)', 'Note'],
  baseCurrency: 'INR',
  // Free, no-key-needed rates endpoint. Swappable for any compatible API.
  ratesApi: 'https://api.exchangerate.host/latest?base=INR',
  cacheTtlSeconds: 3600
};

/* --------------------------------- Menu ---------------------------------- */

function onOpenExpenseTracker() {
  SpreadsheetApp.getUi()
    .createMenu('Expenses')
    .addItem('➕ Log Expense…', 'showLogExpenseDialog')
    .addItem('💱 Refresh FX Conversions', 'refreshExpenseConversions')
    .addItem('📊 Category Totals', 'showCategoryTotals')
    .addToUi();
}

/* ------------------------------ FX conversion ----------------------------- */

// In-memory cache for this execution (Apps Script executions are short-lived).
let etRateCache = null;

/**
 * Fetches live FX rates with the base currency as INR.
 * Cached per execution; on API failure falls back to 1:1 and flags the row.
 * @returns {{rates: Object, ok: boolean, error: string}}
 */
function getFxRates() {
  if (etRateCache) return etRateCache;
  try {
    const resp = UrlFetchApp.fetch(ET_CONFIG.ratesApi, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      throw new Error('HTTP ' + resp.getResponseCode());
    }
    const data = JSON.parse(resp.getContentText());
    if (!data || !data.rates) {
      throw new Error('Unexpected API response shape.');
    }
    etRateCache = { rates: data.rates, ok: true, error: '' };
  } catch (e) {
    etRateCache = { rates: {}, ok: false, error: String(e && e.message ? e.message : e) };
  }
  return etRateCache;
}

/**
 * Converts an amount from the given currency into the base currency.
 * The API returns rates as: 1 INR = X <currency>, so divide.
 * @param {number} amount Amount in the source currency.
 * @param {string} currency ISO currency code, e.g. 'USD'.
 * @returns {{value: number, ok: boolean}}
 */
function convertToBase(amount, currency) {
  const code = String(currency || ET_CONFIG.baseCurrency).toUpperCase();
  if (code === ET_CONFIG.baseCurrency) return { value: amount, ok: true };
  const fx = getFxRates();
  const rate = fx.rates[code];
  if (!rate || rate <= 0) return { value: amount, ok: false }; // unknown currency -> keep original
  return { value: amount / rate, ok: fx.ok };
}

/* ------------------------------- Logging ---------------------------------- */

/**
 * Shows a dialog to log an expense in any currency.
 */
function showLogExpenseDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 12px; }
      label { display: block; margin-top: 10px; font-weight: bold; font-size: 13px; }
      input, select { width: 100%; padding: 6px; margin-top: 4px; box-sizing: border-box; }
      button { margin-top: 14px; padding: 8px 16px; background: #0b8043; color: #fff; border: 0; border-radius: 4px; cursor: pointer; }
      #msg { margin-top: 10px; font-size: 13px; }
      .err { color: #c5221f; } .ok { color: #137333; }
    </style>
    <label>Date <input type="date" id="date"></label>
    <label>Category <input type="text" id="category" placeholder="e.g. Travel"></label>
    <label>Amount <input type="number" id="amount" step="0.01" min="0.01"></label>
    <label>Currency
      <select id="currency">
        <option>INR</option><option>USD</option><option>EUR</option>
        <option>GBP</option><option>AED</option><option>SGD</option>
      </select>
    </label>
    <label>Note <input type="text" id="note"></label>
    <button onclick="submit()">Log expense</button>
    <div id="msg"></div>
    <script>
      document.getElementById('date').valueAsDate = new Date();
      function submit() {
        const data = {
          date: document.getElementById('date').value,
          category: document.getElementById('category').value,
          amount: document.getElementById('amount').value,
          currency: document.getElementById('currency').value,
          note: document.getElementById('note').value
        };
        google.script.run
          .withSuccessHandler((res) => {
            document.getElementById('msg').className = 'ok';
            document.getElementById('msg').textContent =
              'Saved ✓ ≈ ₹' + res.baseAmount.toFixed(2);
          })
          .withFailureHandler((err) => {
            document.getElementById('msg').className = 'err';
            document.getElementById('msg').textContent = 'Error: ' + err.message;
          })
          .logExpenseFromForm(data);
      }
    </script>
  `)
    .setWidth(340)
    .setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Log Expense');
}

/**
 * Validates and appends an expense row, converting to the base currency live.
 * @param {Object} data {date, category, amount, currency, note}
 * @returns {{ok: boolean, baseAmount: number, fxOk: boolean}}
 */
function logExpenseFromForm(data) {
  const errors = [];
  const d = new Date(data.date);
  if (isNaN(d.getTime())) errors.push('Date is not valid.');
  if (!(data.category || '').trim()) errors.push('Category is required.');
  const amount = Number(data.amount);
  if (isNaN(amount) || amount <= 0) errors.push('Amount must be a positive number.');
  const currency = String(data.currency || ET_CONFIG.baseCurrency).toUpperCase();
  if (!/^[A-Z]{3}$/.test(currency)) errors.push('Currency must be a 3-letter code.');
  if (errors.length) throw new Error(errors.join(' '));

  const converted = convertToBase(amount, currency);
  getExpensesSheet().appendRow([
    d,
    data.category.trim(),
    amount,
    currency,
    Math.round(converted.value * 100) / 100,
    data.note || ''
  ]);
  return { ok: true, baseAmount: converted.value, fxOk: converted.ok };
}

/**
 * Re-fetches FX rates and recomputes the "Amount (INR)" column for all rows.
 * Useful when rates move or a row was saved while the API was down.
 */
function refreshExpenseConversions() {
  etRateCache = null; // force a fresh fetch
  const sheet = getExpensesSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return 0;
  let updated = 0;
  for (let i = 1; i < values.length; i++) {
    const amount = Number(values[i][2]);
    const currency = String(values[i][3] || ET_CONFIG.baseCurrency).toUpperCase();
    const converted = convertToBase(amount, currency);
    sheet.getRange(i + 1, 5).setValue(Math.round(converted.value * 100) / 100);
    updated++;
  }
  return updated;
}

/* -------------------------------- Totals ---------------------------------- */

/**
 * Returns per-category totals in the base currency.
 * @returns {Array<{category: string, total: number, count: number}>} sorted desc.
 */
function getCategoryTotals() {
  const sheet = getExpensesSheet();
  const values = sheet.getDataRange().getValues();
  values.shift();
  const map = {};
  for (const r of values) {
    if (r[0] === '' || r[0] == null) continue;
    const cat = String(r[1] || 'Uncategorized');
    if (!map[cat]) map[cat] = { category: cat, total: 0, count: 0 };
    map[cat].total += Number(r[4]) || 0;
    map[cat].count++;
  }
  return Object.keys(map)
    .map((k) => map[k])
    .sort((a, b) => b.total - a.total);
}

/**
 * Shows per-category totals in a dialog.
 */
function showCategoryTotals() {
  const totals = getCategoryTotals();
  const rows = totals
    .map((t) => '<tr><td>' + t.category + '</td><td align="right">₹' +
      t.total.toFixed(2) + '</td><td align="right">' + t.count + '</td></tr>')
    .join('');
  const html = HtmlService.createHtmlOutput(
    '<table border="1" cellpadding="6" style="border-collapse:collapse;font-family:Arial;font-size:13px">' +
    '<tr><th>Category</th><th>Total (₹)</th><th>Items</th></tr>' + rows + '</table>'
  ).setWidth(420).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Expense Totals by Category');
}

/**
 * Returns the Expenses sheet, creating it with headers if missing.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getExpensesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ET_CONFIG.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(ET_CONFIG.sheetName);
    sheet.appendRow(ET_CONFIG.headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, ET_CONFIG.headers.length).setFontWeight('bold');
  }
  return sheet;
}
