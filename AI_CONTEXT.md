# Project Context: Ice Cream Dreams Finance Agent

Use this document as a portable briefing when asking an AI assistant to work on this repository. It describes the checked-in implementation, not a fully reconciled accounting system. Confirm current code and configuration before making changes.

`PROJECT_REQUIREMENTS.md` records the target product requirements and an implementation coverage matrix. Treat those requirements as goals, not as evidence that the current code already supports them.

# Project Context: Ice Cream Dreams Finance Agent

Use this document as a portable briefing when asking an AI assistant to work on this repository. It describes the checked-in implementation, not a fully reconciled accounting system. Confirm current code and configuration before making changes.
## Purpose

This is a local Python application that reads the configured Excel workbook in read-only mode by default, validates and summarizes transactions, and displays results in a Streamlit dashboard. Google Sheets remains an optional source. A separate process records a scheduled daily refresh. The project is an operational baseline for expense monitoring, not an accounting, tax, revenue, profit, or cash-balance system.

## Repository Map
 `excel_source.py`: read-only loading of workbook worksheet values from the configured Excel file.
 `google_sheets.py`: optional Google OAuth setup and read-only retrieval of spreadsheet metadata and worksheet values.
- `google_sheets.py`: Google OAuth setup and read-only retrieval of spreadsheet metadata and worksheet values.
- `processor.py`: worksheet/header detection, column mapping, parsing, category normalization, validation, period metrics, and SQLite run history.
The local source workbook `Ice Cream Dreams - Expense Report.xlsx` is financial data and is excluded from Git by spreadsheet ignore rules. Its `Monhtly_Expense` tab has a row 3 header matching the expected transaction fields and is configured as the active ledger. Other transaction-like tabs are not combined automatically.
- `app.py`: Streamlit dashboard, refresh behavior, filters, charts, transaction table, CSV download, and issue display.
- `scheduler.py`: blocking APScheduler process for the daily refresh.
- `README.md`: installation, setup, operations, troubleshooting, and documented limitations.
2. `excel_source.read_workbook()` opens `WORKBOOK_PATH` using `openpyxl` in read-only, cached-value mode and returns worksheet rows. Set `SOURCE_TYPE=google_sheets` to use the optional OAuth reader instead.
2. `processor.detect_ledger_sheet()` selects `LEDGER_SHEET` when configured; otherwise it detects candidate tabs from their headers. It scans the first 20 rows for a likely header and ignores blank cells and rows.
- `com.icecreamdreams.financeagent.plist.example`: macOS LaunchAgent template; paths must be customized before use.
- `.github/copilot-instructions.md`: repository-wide guidance to keep this context current after AI-assisted code/configuration changes.
5. `run_refresh()` returns the processed DataFrame, issue DataFrame, metrics, selected worksheet, source type/location, and source title. Scheduled/one-shot refreshes record run metadata and metrics in SQLite; dashboard refreshes pass `record_run=False`.
- `.github/agents/context-maintainer.agent.md`: dedicated agent for auditing and updating this context on request.
- `PROJECT_REQUIREMENTS.md`: target product requirements, current coverage, and open workbook-dependent inputs.
 `SOURCE_TYPE`: defaults to `excel`; set to `google_sheets` for the optional Google Sheets reader.
 `WORKBOOK_PATH`: defaults to `Ice Cream Dreams - Expense Report.xlsx` beside the application; absolute and project-relative paths are supported.
 `LEDGER_SHEET`: defaults in `.env.example` to `Monhtly_Expense`; otherwise auto-detection is used.
 `SPREADSHEET_ID`: required only for optional Google Sheets mode; do not commit a live identifier.

 `SCHEDULE_TIME`: daily scheduler time, default `09:00` in `TIMEZONE`.

The default Excel source requires no OAuth. Optional Google Sheets mode requires local Desktop OAuth files named `credentials.json` and `token.json`. These, `.env`, SQLite databases, logs, and Excel workbooks are excluded from Git. Never ask a user to paste OAuth secrets or tokens into an AI conversation.

With Python 3.10.7+, the workbook in the project directory, and a configured `.env` if needed:
2. `processor.detect_ledger_sheet()` selects `LEDGER_SHEET` when configured; otherwise it detects candidate tabs from their headers. It scans the first 20 rows for a likely header and reads values from columns A through Z.
 Keep both Excel and optional Google Sheets source access read-only.
4. `calculate_metrics()` computes lifetime, daily, weekly, monthly, previous-month, and financial-year expense metrics using `TIMEZONE` and `FINANCIAL_YEAR_START_MONTH`.
 The `Consolidated` formulas have been inspected structurally but not independently recalculated or reconciled against the ledger.
 Other transaction-like workbook tabs are not automatically included; `Monhtly_Expense` is the configured ledger.
 Excel dates and unusual merged/multi-row headers may require mapping adjustments.

## Data Contract and Rules

The expected ledger fields are Date, Invoice, Type/category, Particular/description, Amount (INR), Paid By, and Comments. Source column names are normalized to lowercase with spaces converted to underscores and periods removed. The processor maps common variants such as `INR` and `amount_inr` to the amount field, and preserves canonical transaction fields such as `amount_inr` and `comments`. Column aliases and category rules are defined in `processor.py` and should be extended conservatively when the real source headings or categories require it.

Date and amount are required for calculation. A row with either missing or invalid value remains in the displayed/validated data but is excluded from period totals. Negative amounts remain in totals and are flagged for review; do not assume whether they are refunds or adjustments. Duplicate invoice numbers and potential duplicate rows are warnings, not automatic deletions. Missing descriptions/categories, unmapped categories, and optionally high amounts are flagged for review.

The `Consolidated` tab may be identified, but no formula-by-formula reconciliation is implemented. Do not claim that ledger totals have been reconciled to that tab. Category classification is deterministic; there is no external LLM or AI service in the current implementation.

## Configuration

Configuration is loaded from `.env` when present. `.env.example` documents:

- `SPREADSHEET_ID`: ID of the source Google spreadsheet; treat as deployment-specific configuration.
- `LEDGER_SHEET`: optional exact worksheet tab name; otherwise auto-detection is used.
- `TIMEZONE`: defaults to `Asia/Kolkata`.
- `SCHEDULE_TIME`: daily scheduler time, default `14:00` in `TIMEZONE`.
- `DASHBOARD_REFRESH_SECONDS`: dashboard refresh interval, default `300` seconds while the page is open.
- `FINANCIAL_YEAR_START_MONTH`: defaults to `4` (April).
- `HIGH_EXPENSE_THRESHOLD`: optional INR warning threshold; blank disables it.

Google OAuth requires a local Desktop OAuth client file named `credentials.json`. The authorized token is stored in `token.json`. Both files are sensitive and are ignored by Git along with `.env`, SQLite databases, and logs. Never ask a user to paste their OAuth secrets or token into an AI conversation.

## Running Locally

With Python 3.10.7+ and a configured `.env`, `credentials.json`, and spreadsheet access:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python agent.py
streamlit run app.py
python scheduler.py
```

The dashboard and scheduler are separate processes. The scheduler must remain running (or be managed by a process manager) to execute daily. The README contains the fuller macOS LaunchAgent setup.

## Safety and Change Guidance

- Preserve the read-only Google Sheets scope and behavior unless the user explicitly authorizes a change to source data access.
- Treat this as financial data: avoid exposing transaction details, spreadsheet identifiers, OAuth material, local database contents, or logs unnecessarily.
- Do not infer revenue, profit, balances, payment status, refund semantics, or accounting reconciliation from expense rows alone.
- Keep validation findings distinct from corrections. Never silently delete or rewrite source transactions.
- For changes to parsing, category rules, or metrics, use representative edge cases (ambiguous dates, formatted currency, negative values, missing fields, duplicates, and timezone/period boundaries).
- Follow the existing Python style and keep changes scoped. Check available tests and diagnostics before claiming behavior is verified; this repository currently documents no dedicated test suite.
- Read `README.md` and the relevant implementation before changing behavior. Treat comments and this context as a guide, and verify against current code.

## Known Limitations

- The source workbook's `Consolidated` formulas have not been inspected or reconciled with the ledger.
- Google Sheets date formatting and unusual multi-row or merged headers may need mapping adjustments.
- Amount/date interpretation and category mapping are intentionally heuristic/conservative.
- Dashboard refresh occurs only while its page is open; the daily scheduler is independent.
- Run history stores metadata and metrics, not a transaction-level SQLite mirror.
- The current implementation does not send notifications.