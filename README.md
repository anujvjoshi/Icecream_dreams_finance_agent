# Ice Cream Dreams – Expense Intelligence Agent

A local Python application that reads `Ice Cream Dreams - Expense Report.xlsx` in read-only mode, validates transactions, calculates expense summaries, and presents a Streamlit dashboard. A separate scheduler runs the daily refresh at 09:00 Asia/Kolkata.

## 1. Requirements

- macOS with Python 3.10.7 or newer
- VS Code
- The supplied Excel workbook in this project folder, or a configured workbook path
- Optional Google account and OAuth client only when using Google Sheets mode

## 2. Configure the project

Create a virtual environment and install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Create `.env` from the example:

```bash
cp .env.example .env
```

The default configuration reads `Ice Cream Dreams - Expense Report.xlsx` from this project directory and selects the `Monhtly_Expense` worksheet. Set `WORKBOOK_PATH` to an absolute path or a path relative to the project directory to use another workbook. Set `LEDGER_SHEET` to the exact transaction tab name if needed.

The source is read-only. The workbook is excluded from Git; do not remove its ignore rule or commit financial source data.

## 3. First test run

Run:

```bash
python agent.py
```

The agent reads the configured workbook and prints the detected ledger tab, transaction count, total and validation issue count. No OAuth setup is required for the default Excel source.

## 5. Start the dashboard

In a VS Code terminal:

```bash
streamlit run app.py
```

Streamlit will print a local URL, normally:

`http://localhost:8501`

The dashboard reads the source when it starts, refreshes automatically every 5 minutes while the browser page is open, and has a manual refresh button. It also shows the last successful dashboard refresh time. Excel is the default source; `SOURCE_TYPE=google_sheets` enables the optional Google Sheets reader.

## 6. Start the daily scheduler

Open a second VS Code terminal, activate the same virtual environment, and run:

```bash
python scheduler.py
```

The scheduler runs every day at **09:00 Asia/Kolkata** by default. Keep this terminal/process running. The scheduler is separate from Streamlit.

For reliable unattended operation, run the scheduler as a macOS LaunchAgent or another process manager. The Mac must be powered on, awake, connected to the internet, and logged in/authorized as needed.

A LaunchAgent template is included as `com.icecreamdreams.financeagent.plist.example`. Replace every `/ABSOLUTE/PATH/TO/PROJECT` with the actual project folder path, save the finished file as `com.icecreamdreams.financeagent.plist` under `~/Library/LaunchAgents/`, then load it:

```bash
launchctl load ~/Library/LaunchAgents/com.icecreamdreams.financeagent.plist
launchctl start com.icecreamdreams.financeagent
```

To stop and unload it:

```bash
launchctl unload ~/Library/LaunchAgents/com.icecreamdreams.financeagent.plist
```

## 7. Run both processes in VS Code

Terminal 1:

```bash
source .venv/bin/activate
streamlit run app.py
```

Terminal 2:

```bash
source .venv/bin/activate
python scheduler.py
```

## 8. How the data is handled

- The Excel source is opened in read-only mode; the source workbook is never modified.
- The app detects the ledger worksheet from its column headers; `LEDGER_SHEET` can select an exact tab.
- The dashboard compares `Monhtly_Expense` amounts with Kiosk/DNS cash and online receipts from `DAILY_TALLY`. Margin is provisional; attribution/platform income, investments, and Prasad account movements are shown separately to avoid unverified double counting.
- The dashboard provides an overview and on-demand row explorer for every workbook sheet. Values are read from stored/cached cells; formulas are not recalculated by the app.
- Investment/vendor summary includes each source/payer group's percentage of the total displayed Investment/Delrico amount; it is separate from operating expenses.
- The `Consolidated` worksheet has been structurally inspected, but its formulas are not independently reconciled against the transaction ledger. Do not claim reconciliation.
- The app flags missing/invalid dates and amounts, missing descriptions/categories, duplicate invoice numbers, potential duplicate rows and negative amounts.
- Rows with invalid dates or amounts are excluded from period totals and remain visible as validation issues.
- Negative values remain in totals but are flagged for review; confirm whether they are refunds/adjustments.
- Category normalization is intentionally conservative and configurable in `processor.py`.
- The SQLite database stores run metadata and metric snapshots. It does not modify the source workbook.

## 9. Troubleshooting

### Workbook not found
Set `WORKBOOK_PATH` in `.env` to the workbook's absolute path or a path relative to the project directory.

### Google Sheets access (optional mode)
Set `SOURCE_TYPE=google_sheets`, provide `SPREADSHEET_ID`, and configure a Desktop OAuth client as described in the Google Sheets API Python quickstart. Keep `credentials.json` and `token.json` private.

### Wrong ledger worksheet selected
Set `LEDGER_SHEET=Exact Tab Name` in `.env`, then restart the app/agent.

### No ledger detected
Check that a header row contains recognizable fields such as Date, Invoice, Type, Particular, Amount (INR), Paid By and Comments. Add exact aliases in `ALIASES` inside `processor.py` if your actual headings differ.

### Scheduler did not run
Confirm `python scheduler.py` is still running, the Mac is awake, and the configured timezone is `Asia/Kolkata`.

## 10. Important limitations of this starter implementation

1. It is a working baseline for Excel source access, validation, aggregation and dashboarding, but it is not yet a fully reconciled finance product.
2. The `Consolidated` tab has been inspected structurally; exact formula-by-formula reconciliation is not implemented.
3. AI narrative generation is deterministic and based on calculated data; no external LLM/API key is required.
4. Excel dates and unusual merged/multi-row headers may require a small mapping adjustment.
5. The app does not infer revenue, profit, payment due status or cash balance from expense data alone.
6. The scheduler records successful/failed runs in SQLite and logs. It does not send email/WhatsApp notifications in this version.

## 11. Security

Never commit the source workbook, `credentials.json`, `token.json`, `.env`, SQLite databases, or logs to Git. The workbook contains financial data. Google OAuth files are needed only for optional Google Sheets mode and must also remain private.
