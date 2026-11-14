# Project Context: Ice Cream Dreams Finance Agent

Use this as a portable briefing for AI assistants working in this repository. It describes the current implementation, not a fully reconciled finance system. `PROJECT_REQUIREMENTS.md` records target requirements and factual coverage statuses.

## Purpose

This local Python application reads the configured Excel workbook in read-only mode by default, validates and summarizes expense transactions, and displays results in a Streamlit dashboard. Google Sheets remains an optional source. The app does not calculate verified profit, revenue completeness, or cash balances.

## Repository Map

- `excel_source.py`: read-only loading of workbook worksheet values using `openpyxl`.
- `google_sheets.py`: optional Google OAuth and read-only Sheets API source.
- `processor.py`: sheet/header detection, parsing, validation, monthly metrics, and SQLite run history.
- `app.py`: Streamlit dashboard, finance comparisons, worksheet explorer, filters, charts, and CSV export.
- `agent.py`: one-shot refresh and execution logging.
- `scheduler.py`: daily APScheduler process.
- `README.md`: setup, operations, and limitations.
- `PROJECT_REQUIREMENTS.md`: target requirements and current coverage.
- `.github/copilot-instructions.md` and `.github/agents/context-maintainer.agent.md`: context maintenance guidance and agent.

## Workbook and Data Flow

The local source is `Ice Cream Dreams - Expense Report.xlsx`, excluded from Git. The workbook contains seven sheets. `Monhtly_Expense` is configured as the active operating ledger; other transaction-like sheets are not silently merged into that ledger.

1. `excel_source.read_workbook()` resolves `WORKBOOK_PATH`, opens the workbook with `read_only=True` and `data_only=True`, and returns values for every worksheet without modifying the source.
2. `processor.run_refresh()` defaults to `SOURCE_TYPE=excel`; `SOURCE_TYPE=google_sheets` opts into the legacy OAuth reader.
3. `detect_ledger_sheet()` selects `LEDGER_SHEET` when configured or detects a candidate from headers in the first 20 rows. Blank `None` and NaN cells/rows are ignored.
4. `prepare_transactions()` parses `Monhtly_Expense`, normalizes categories, and flags issues. Invalid dates/amounts remain visible but are excluded from expense metrics.
5. `prepare_daily_income()` uses Kiosk/DNS cash and online columns from `DAILY_TALLY` to calculate reported receipts. Separate attributed/platform income fields are displayed but excluded from margin because overlap is unverified.
6. Investment/Delrico spending and Prasad received/spent activity are parsed and displayed separately; they are not added to operating-ledger expense or receipt totals. Investment/vendor spend shares use the total displayed Investment/Delrico amount as the denominator. The dashboard includes an overview and on-demand row explorer for every sheet.
7. Monthly comparison uses reported Kiosk/DNS receipts versus amounts in `Monhtly_Expense`. Margin is `(receipts - ledger amount) / receipts`; month-to-month margin movement is the percentage-point difference from the previous calendar month. Treat this as provisional, not verified profit margin.

A read-only check on the supplied workbook found 125 calculable ledger transactions, 333 dated receipt rows, 135 dated investment/vendor records, 38 Prasad activity rows, and 13 monthly periods. These are structural/runtime check counts, not a claim of reconciliation.

## Configuration

Configuration is loaded from `.env` when present. `.env.example` documents:

- `SOURCE_TYPE`: defaults to `excel`; set to `google_sheets` for the optional Sheets reader.
- `WORKBOOK_PATH`: defaults to `Ice Cream Dreams - Expense Report.xlsx`; supports absolute or project-relative paths.
- `LEDGER_SHEET`: defaults to `Monhtly_Expense` in `.env.example`.
- `TIMEZONE`: defaults to `Asia/Kolkata`.
- `SCHEDULE_TIME`: scheduler default is `09:00`; local `.env` can override it.
- `DASHBOARD_REFRESH_SECONDS`: defaults to 300 seconds while the dashboard is open.
- `FINANCIAL_YEAR_START_MONTH`: defaults to 4 (April).
- `HIGH_EXPENSE_THRESHOLD`: optional INR review threshold.
- `SPREADSHEET_ID`: needed only for optional Google Sheets mode; do not commit a live identifier.

The default Excel source requires no OAuth. Optional Google Sheets mode requires private local `credentials.json` and `token.json`. `.env`, Excel workbooks, credentials, local databases, and logs must remain untracked and must not be copied into AI prompts.

## Running

With Python 3.10.7+:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python agent.py
streamlit run app.py
python scheduler.py
```

The dashboard and daily scheduler are separate processes. See `README.md` for LaunchAgent setup.

## Safety and Limitations

- Keep both source adapters read-only. Never rewrite, delete, or silently correct source transactions.
- Treat workbook data, spreadsheet identifiers, OAuth material, local databases, and logs as sensitive.
- Duplicate and missing-data findings are review flags, not automatic corrections. Negative amounts remain in totals and require interpretation.
- The displayed margin is provisional: it excludes unverified attributed/platform receipts, and transfers or ledger classifications may distort the expense side.
- `Consolidated` has been structurally inspected, but its formulas have not been independently recalculated or reconciled against the ledger. Do not claim reconciliation.
- Workbook modification monitoring, last-successful dashboard snapshot recovery, notifications, and transaction-level history are not implemented.
- There is no dedicated test suite in the repository. Validate parsing changes against representative cases and do not claim unrun acceptance checks.
