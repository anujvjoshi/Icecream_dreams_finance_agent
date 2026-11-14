# Product Requirements and Coverage

This document records the requested target from the Ice Cream Dreams master development brief and distinguishes it from the current implementation. It is a requirements baseline, not a claim that every target feature is already delivered.

## Source of Truth and Open Inputs

The target system must read the existing `Ice Cream Dreams - Expense Report.xlsx` workbook as its source of truth, in read-only mode. This workbook is now present in the project directory. Its sheets include `Monhtly_Expense`, `DAILY_TALLY `, `Investment`, `Consolidated`, `Prasad`, `Delrico`, and `Sheet9`. `Monhtly_Expense` row 3 has headers matching Date, Invoice, Type, Particular, INR, Paid By, and Comment, making it the strongest primary-ledger candidate; confirm before treating it as authoritative because other tabs may contain additional transaction sources.

The application now reads the configured Excel workbook in read-only mode by default; Google Sheets remains an explicit optional source. `Monhtly_Expense` is configured as the active transaction tab. The workbook path is configurable, but modification monitoring and persisted last-successful dashboard data are not implemented.

## Coverage Status

Statuses describe the code currently in the repository, not the desired final state. Recheck source files before changing these labels.

| Requirement area | Current status | Evidence or gap |
| --- | --- | --- |
| Read-only source access | Implemented for selected Excel source | `openpyxl` reads the configured workbook in read-only, cached-value mode. Google Sheets remains optional. |
| Workbook monitoring | Not implemented | No modification-time check, change detection, transaction diffing, structural-change report, or unchanged-source skip. |
| Header-based transaction extraction | Partial | Header detection and aliases read the configured `Monhtly_Expense` sheet; 125 calculable transactions were verified in a local refresh. Other transaction-like tabs are not combined, and arbitrary original fields are not all preserved. |
| Transaction validation | Partial | Dates, amounts, descriptions/categories, duplicate invoices/rows, negatives, and optional high amounts are flagged. Missing invoice and inconsistent payment-description checks are absent. Issues do not provide corrective actions. |
| Category normalization | Partial | Rules are hardcoded in `processor.py`; the original type field is retained beside a normalized category. Mapping is not separately configurable/auditable, and uncertain Snowball classification is heuristic. |
| Consolidated analysis/reconciliation | Not implemented; partially inspected | The workbook's `Consolidated` tab contains 10 formulas for outflow/inflow totals, order value, person-level balances, a balance total, and an expected total. No cached formula-error cells were found. Formula groupings and calculated outputs have not been reconciled against ledger transactions. |
| Financial metrics | Partial | Lifetime, today, yesterday, week, month, previous month, month change, month daily average, financial year, counts, and extrema are calculated. Payment-type metrics and category/payer count, share, and average metrics are absent. |
| Dashboard | Partial | KPI cards, category bar/donut charts, monthly trend, payer/category/date filters, inventory summary, issue table, and CSV export exist. Daily trend/quick ranges, search, Excel export, complete category/payer statistics, and several requested KPIs/analyses are absent. |
| Inventory and operational analysis | Partial | Inventory/Snowball totals and transactions are shown. Monthly inventory trend, largest purchases, recurring-cost analysis, and cautious uncertain-classification workflow are absent. |
| AI management insights | Not implemented in dashboard | No LLM integration or traceable management-summary module is present. Do not invent causes, revenue, profit, balances, or other unsupported business facts. |
| Daily automation | Partial | APScheduler defaults to 09:00 in `TIMEZONE` and logs runs. It does not detect unchanged workbooks or send notifications. |
| Failure continuity and history | Partial | SQLite stores run metadata and metric snapshots, not transaction-level snapshots. The app reports refresh errors, but does not render a persisted last-successful dashboard snapshot after a failed live refresh. |
| Data protection | Partial/current | Excel and Google Sheets adapters are read-only; the workbook, credentials, local config, database, and logs are ignored by Git. The public repository requires no live spreadsheet identifier in tracked files. |
| Tests and acceptance report | Partial / no suite | No dedicated test suite is present. A read-only ingestion check ran on the actual workbook; independent Consolidated reconciliation and the full acceptance suite remain outstanding. |

## Target Behavior

### Data Access and Integrity

- Read the actual configured `.xlsx` source; do not use a hardcoded or manually exported substitute. Report missing, locked, inaccessible, or corrupt files and retain the last successful results.
- Identify worksheets and transaction columns from headers, not fixed column positions. Monitor workbook modification time and detect added, updated, deleted, or structurally changed transactions without double-counting.
- Preserve original transaction values alongside normalized analytical fields. Never modify the source workbook, transactions, formulas, or Consolidated sheet.
- Analyze the complete Consolidated sheet and its formulas as a reference, independently recalculate metrics from the ledger, and report discrepancies. Never claim reconciliation until it is verified against the actual workbook.

### Validation and Financial Logic

- Flag missing/invalid dates and amounts, missing invoice/category/particular, duplicates, negative amounts, configurable unusual/high expenses, inconsistent payment descriptions, and unrecognized categories. Never silently delete or rewrite a questionable transaction.
- Keep category mappings configurable and auditable; discover and retain categories beyond the initial known list. Preserve original category text and make uncertain classifications reviewable.
- Calculate lifetime, today, yesterday, week, month, prior month, month-over-month change, financial year (configurable; default April to March), average daily spend, highest/lowest transactions, category totals, payer totals, payment-type totals when available, and transaction/issue counts.
- Keep recorded, calculable/validated, and review-required amounts distinct. Keep unclassified expenses visible. Define denominators and reporting periods for all percentages and averages. Treat negative values as adjustments requiring review unless verified otherwise.
- Do not infer revenue, profit, cash balance, reimbursement status, or causal explanations that the source does not support.

### Dashboard and Reporting

- Provide a responsive management dashboard with lifetime/today/current-month/prior-month/average-daily/highest-expense/transaction-count/issue-count KPIs, INR formatting, and an explicit last-successful refresh time.
- Include category totals, share, counts and averages; monthly and daily trends; date-range shortcuts and custom filtering; payer statistics; inventory/Snowball analysis; fixed/operational cost analysis; searchable transaction ledger; and filtered CSV/Excel export.
- Show validation issues and Consolidated discrepancies with affected transaction references and review guidance. Do not label incomplete data as reconciled.
- Generate management insights only from calculated, traceable figures and transactions. Clearly distinguish deterministic calculations from any AI-generated narrative.

### Automation, History, and Acceptance

- Target daily execution is 09:00 Asia/Kolkata unless configured otherwise. Verify the source even when it appears unchanged; avoid unnecessary duplicate processing while keeping the dashboard current.
- Record execution status, refresh time, and sufficient historical data to support reporting. Notify on important issues when a notification channel is configured. A failed run must not destroy or replace the last successful dashboard data.
- Test actual workbook extraction, independent reconciliation, totals and category rollups, duplicates/missing values, monthly calculations, filters/exports, added/updated transactions, unchanged-source behavior, source immutability, failure recovery, and scheduling. Report expected versus actual results; mark workbook-dependent checks blocked when the real source is unavailable.

## Implementation Guidance

- Before implementing a requirement, inspect the current owning modules and update this table based on observed behavior.
- Keep target requirements separate from delivered functionality. Mark a requirement covered only after implementation and an appropriate check; document partial support and blockers.
- Use `Monhtly_Expense` as the configured ledger unless the user changes that source selection. Do not claim reconciliation until Consolidated formula outputs have been independently recalculated and compared with ledger results.
- Keep the data layer read-only. Do not add a remote LLM or upload workbook/transaction content without explicit user authorization and a configured, approved integration.