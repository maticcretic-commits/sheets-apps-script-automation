# Sheets + Apps Script Automation

**Google Sheets + Apps Script automation — a transaction management system and multi-currency expense tracker with live dashboards.**

## The problem

Small businesses run on spreadsheets — and drown in them. Transactions get pasted in by hand, categories are inconsistent, nobody knows this month's real profit until someone spends a weekend reconciling, and expenses paid in USD/AED/GBP sit in the sheet unconverted, so totals are fiction. Hiring an accountant is overkill; the manual spreadsheet grind is not.

This project replaces that grind with an Apps Script layer inside the spreadsheet itself: validated entry forms, automatic categorization, live currency conversion, and a dashboard that rebuilds itself with one click.

## Features

- **Transaction management system** (`TransactionManager.gs`)
  - Custom "Transactions" menu in the sheet: Add Transaction, Monthly Summary, one-click sheet formatting.
  - Dialog-based entry form (no touching raw cells), with validation: type must be Income/Expense, date must be valid and not in the future, amount must be positive.
  - **Auto-categorization**: leave the category blank and the script infers it from keywords in the note ("Uber" → Transport, "rent" → Rent, …). Rules are a plain config list, easy to extend.
  - Monthly aggregation (income / expense / net, per-category breakdown).
- **Multi-currency expense tracker** (`ExpenseTracker.gs`)
  - Log expenses in any currency (USD, EUR, GBP, AED, SGD, …).
  - **Live FX conversion** via a free exchange-rates API, cached per execution; every row stores the converted base-currency amount.
  - One-click "Refresh FX Conversions" re-prices all rows when rates move; per-category totals dialog.
- **Live dashboard** (`Dashboard.gs`)
  - Rebuilds a Dashboard sheet on demand: headline KPIs, this month's summary table, expense categories table, alerts block.
  - **EmbeddedChart**: pie chart of spend by category + column chart of income vs expenses.
  - Conditional formatting: negative net in red, positive in green; over-threshold amounts highlighted.
  - **Alerts**: flags categories over a configurable high-spend threshold and negative monthly net.
- **Ready-to-import templates** (`sheets-templates/`) — CSVs with headers + 5 sample rows each, so the system runs end-to-end in minutes.
- **Sanity-checked** (`tests/validate.py`) — every `.gs` file passes `node --check`; templates parse with exact headers.

## Quickstart

1. **Create the spreadsheet.** Open [Google Sheets](https://sheets.google.com) and make a blank spreadsheet (e.g. "Business Finance").
2. **Import the templates.** File → Import → Upload the two CSVs from `sheets-templates/` (choose "Create new sheet" for each), or just copy the headers into `Transactions` and `Expenses` sheets manually.
3. **Open Apps Script.** Extensions → Apps Script. Delete the default `Code.gs`.
4. **Paste the scripts.** Create three files — `TransactionManager.gs`, `ExpenseTracker.gs`, `Dashboard.gs` — and paste in the contents from `apps-script/`.
5. **Authorize once.** Run `onOpen()` (or `setupMenu()`) and grant the permissions it asks for.
6. **Use the menus.** Reload the sheet → new **Transactions** and **Expenses** menus appear. Add a transaction, log a USD expense, then Transactions → Generate Monthly Summary.

## Sheet layout

```
┌─ Transactions ───────────────────────────────────────────────────┐
│ Date       │ Type    │ Category │ Amount   │ Note        │ CreatedBy │
│ 2026-09-01 │ Income  │ Revenue  │ 85000.00 │ Client inv. │ you@…     │
│ 2026-09-03 │ Expense │ Rent     │ 22000.00 │ Office rent │ you@…     │
├─ Expenses ───────────────────────────────────────────────────────┤
│ Date       │ Category │ Amount │ Currency │ Amount (INR)│ Note      │
│ 2026-09-02 │ Software │ 49.00  │ USD      │ 4083.66*    │ SaaS sub  │
├─ Dashboard (auto-built) ─────────────────────────────────────────┤
│ 📈 KPI block: Total Income / Total Expenses / Net Balance        │
│ 🗓  Monthly summary table (Type | Category → Amount)             │
│ 💱 Expense categories (base-currency totals + item counts)       │
│ ⚠️  Alerts: high-spend categories, negative net                  │
│ 📊 Embedded pie + column charts                                 │
└──────────────────────────────────────────────────────────────────┘
* converted live via the FX API at log time
```

## Tech stack

- **Google Apps Script** (JavaScript) — SpreadsheetApp, HtmlService, UrlFetchApp, Charts
- **Google Sheets** as the database + UI (menus, dialogs, EmbeddedChart, conditional formatting)
- **exchangerate.host** free API for live FX rates (no key needed; swappable)
- **Node.js** for `.gs` syntax validation in CI/local checks

## Project structure

```
apps-script/
  TransactionManager.gs   # menus, entry dialog, validation, auto-categorization
  ExpenseTracker.gs       # multi-currency logging, live FX conversion, totals
  Dashboard.gs            # KPI/alerts/tables, EmbeddedCharts, formatting
sheets-templates/
  transactions_template.csv
  expenses_template.csv
tests/
  validate.py             # node --check for .gs files + CSV header checks
```

## Running the checks

```bash
python3 tests/validate.py
```

## Notes

- The FX API used requires no key. If it is ever down, rows are still saved (with a flag-worthy fallback) and can be re-priced later with *Refresh FX Conversions*.
- Amounts in the templates' `Amount (INR)` column are `0` on purpose — the script fills them at log time.
