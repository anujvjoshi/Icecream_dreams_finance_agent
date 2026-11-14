import os
import re
import json
import sqlite3
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from dotenv import load_dotenv

from excel_source import get_workbook_path, read_workbook

load_dotenv()
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "finance_agent.db"
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
DB_PATH.parent.mkdir(exist_ok=True)

TIMEZONE = os.getenv("TIMEZONE", "Asia/Kolkata")
REQUIRED = ["date", "invoice", "type", "particular", "amount_inr", "paid_by", "comments"]

ALIASES = {
    "date": ["date", "transaction date", "expense date"],
    "invoice": ["invoice", "invoice no", "invoice number", "bill no", "receipt no"],
    "type": ["type", "category", "expense type", "expense category"],
    "particular": ["particular", "particulars", "description", "item", "details"],
    "amount_inr": ["amount (inr)", "amount inr", "amount", "expense amount", "value"],
    "paid_by": ["paid by", "payer", "payment by", "paidby"],
    "comments": ["comments", "comment", "notes", "remarks"],
}

CATEGORY_RULES = {
    "rent": "Rent",
    "electricity": "Electricity",
    "electric bill": "Electricity",
    "power bill": "Electricity",
    "staff salary": "Staff Salary",
    "salary": "Staff Salary",
    "salaries": "Staff Salary",
    "inventory": "Inventory",
    "stock": "Inventory",
    "snowball": "Snowball Inventory",
    "operational expense": "Operational Expenses",
    "operations": "Operational Expenses",
    "cash purchase": "Cash Purchases",
}

def clean_header(value):
    normalized = str(value or "").strip().lower().replace(".", "").replace("_", " ")
    return re.sub(r"\s+", " ", normalized)

def locate_header_and_frame(rows):
    """Find a likely header row in the first 20 rows."""

    def has_value(value):
        if value is None:
            return False
        if isinstance(value, float) and pd.isna(value):
            return False
        return bool(str(value).strip())

    best = None
    for idx, row in enumerate(rows[:20]):
        normalized = [clean_header(x) for x in row]
        joined = " | ".join(normalized)
        score = sum(
            any(alias in normalized for alias in aliases)
            for aliases in ALIASES.values()
        )
        # Amount and date are essential clues; other columns improve confidence.
        if score >= 4 and (best is None or score > best[0]):
            best = (score, idx, row)
    if best is None:
        return None
    _, header_idx, header = best
    width = len(header)
    data = []
    for row in rows[header_idx + 1:]:
        padded = list(row[:width]) + [""] * max(0, width - len(row))
        if not any(has_value(value) for value in padded):
            continue
        data.append(padded[:width])
    return pd.DataFrame(data, columns=[str(x).strip() for x in header])

def map_columns(df):
    lookup = {clean_header(c): c for c in df.columns}
    mapped = {}
    for canonical, aliases in ALIASES.items():
        for alias in aliases:
            if alias in lookup:
                mapped[canonical] = lookup[alias]
                break
    return mapped

def detect_ledger_sheet(raw_sheets):
    requested = os.getenv("LEDGER_SHEET", "").strip()
    if requested:
        if requested not in raw_sheets:
            raise ValueError(f"Configured LEDGER_SHEET '{requested}' was not found.")
        frame = locate_header_and_frame(raw_sheets[requested])
        if frame is None:
            raise ValueError(f"Could not identify the ledger header in '{requested}'.")
        return requested, frame

    candidates = []
    for name, rows in raw_sheets.items():
        frame = locate_header_and_frame(rows)
        if frame is None:
            continue
        mapped = map_columns(frame)
        score = len(mapped)
        candidates.append((score, name, frame))
    if not candidates:
        raise ValueError(
            "Could not auto-detect an expense ledger. Expected columns such as "
            "Date, Invoice, Type, Particular, Amount (INR), Paid By and Comments."
        )
    candidates.sort(key=lambda x: (x[0], "consolidated" not in x[1].lower()), reverse=True)
    score, name, frame = candidates[0]
    if score < 4:
        raise ValueError("The detected sheet does not contain enough expense columns.")
    return name, frame

def parse_amount(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.replace(",", "").replace("₹", "").replace("INR", "").strip()
    text = text.replace("(", "").replace(")", "")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in ("", "-", ".", "-."):
        return None
    try:
        amount = float(text)
        return -abs(amount) if negative else amount
    except ValueError:
        return None

def normalize_category(value):
    original = str(value or "").strip()
    key = re.sub(r"\s+", " ", original.lower())
    for phrase, category in CATEGORY_RULES.items():
        if phrase in key:
            return category
    return original.title() if original else "Unclassified"

def prepare_transactions(raw_df):
    source_df = raw_df.copy()
    source_df.columns = [
        str(col).strip().lower().replace(".", "").replace(" ", "_")
        for col in source_df.columns
    ]

    column_mapping = {
        "inr": "amount",
        "amount_inr": "amount",
        "comments": "comment",
        "sr_no": "serial_number",
        "sr_no_": "serial_number",
    }
    source_df = source_df.rename(columns=column_mapping)
    mapped = map_columns(source_df)

    required_columns = ["date", "amount"]
    missing_columns = [
        column for column in required_columns
        if ("amount_inr" if column == "amount" else column) not in mapped
    ]
    if missing_columns:
        raise ValueError(
            f"Missing required columns: {missing_columns}. "
            f"Detected columns: {list(source_df.columns)}"
        )

    missing = [c for c in REQUIRED if c not in mapped]

    df = pd.DataFrame()
    for canonical in REQUIRED:
        df[canonical] = source_df[mapped[canonical]] if canonical in mapped else ""

    df["date_original"] = df["date"]
    df["date"] = pd.to_datetime(df["date"], errors="coerce", dayfirst=True)
    df["amount_inr"] = df["amount_inr"].apply(parse_amount)
    for col in ["invoice", "type", "particular", "paid_by", "comments"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    def classify_row(row):
        # Snowball may appear in Particular even when Type is simply Inventory.
        combined = f"{row.get('type', '')} {row.get('particular', '')}".lower()
        if "snowball" in combined:
            return "Snowball Inventory"
        return normalize_category(row.get("type", ""))

    df["category"] = df.apply(classify_row, axis=1)
    df["row_number"] = range(1, len(df) + 1)
    df["duplicate_invoice"] = False
    invoice_mask = df["invoice"].ne("")
    df.loc[invoice_mask, "duplicate_invoice"] = df.loc[invoice_mask, "invoice"].duplicated(keep=False)

    duplicate_cols = ["date", "amount_inr", "particular", "invoice"]
    df["possible_duplicate"] = df.duplicated(subset=duplicate_cols, keep=False)

    issues = []
    for idx, row in df.iterrows():
        row_id = int(row["row_number"])
        if pd.isna(row["date"]):
            issues.append({"row": row_id, "severity": "Error", "issue": "Missing or invalid date", "invoice": row["invoice"], "amount": row["amount_inr"]})
        if row["amount_inr"] is None or pd.isna(row["amount_inr"]):
            issues.append({"row": row_id, "severity": "Error", "issue": "Missing or invalid amount", "invoice": row["invoice"], "amount": None})
        if not row["particular"]:
            issues.append({"row": row_id, "severity": "Warning", "issue": "Missing particular/description", "invoice": row["invoice"], "amount": row["amount_inr"]})
        if not row["type"]:
            issues.append({"row": row_id, "severity": "Warning", "issue": "Missing expense category", "invoice": row["invoice"], "amount": row["amount_inr"]})
        elif row["category"] not in {
            "Rent", "Electricity", "Staff Salary", "Inventory",
            "Snowball Inventory", "Operational Expenses", "Cash Purchases",
            "Other Business Expenses"
        }:
            issues.append({"row": row_id, "severity": "Review", "issue": f"Unmapped category – review mapping: {row['category']}", "invoice": row["invoice"], "amount": row["amount_inr"]})
        if row["duplicate_invoice"]:
            issues.append({"row": row_id, "severity": "Warning", "issue": "Duplicate invoice number", "invoice": row["invoice"], "amount": row["amount_inr"]})
        if row["possible_duplicate"]:
            issues.append({"row": row_id, "severity": "Warning", "issue": "Potential duplicate transaction", "invoice": row["invoice"], "amount": row["amount_inr"]})
        if pd.notna(row["amount_inr"]) and row["amount_inr"] < 0:
            issues.append({"row": row_id, "severity": "Review", "issue": "Negative amount – verify refund/adjustment", "invoice": row["invoice"], "amount": row["amount_inr"]})
        threshold = os.getenv("HIGH_EXPENSE_THRESHOLD", "").strip()
        if threshold and pd.notna(row["amount_inr"]) and row["amount_inr"] > float(threshold):
            issues.append({"row": row_id, "severity": "Review", "issue": "Expense exceeds configured threshold", "invoice": row["invoice"], "amount": row["amount_inr"]})

    # Keep rows with valid dates and amounts for period calculations.
    df["is_calculable"] = df["date"].notna() & df["amount_inr"].notna()
    return df, pd.DataFrame(issues), missing

def calculate_metrics(df):
    valid = df[df["is_calculable"]].copy()
    now = datetime.now(ZoneInfo(TIMEZONE))
    today = pd.Timestamp(now.date())
    month_start = today.replace(day=1)
    week_start = today - pd.Timedelta(days=today.weekday())
    prev_month_end = month_start - pd.Timedelta(days=1)
    prev_month_start = prev_month_end.replace(day=1)
    fy_start_month = int(os.getenv("FINANCIAL_YEAR_START_MONTH", "4"))
    fy_start_year = today.year if today.month >= fy_start_month else today.year - 1
    fy_start = pd.Timestamp(year=fy_start_year, month=fy_start_month, day=1)

    def total(frame):
        return float(frame["amount_inr"].sum()) if not frame.empty else 0.0

    today_df = valid[valid["date"].dt.normalize() == today]
    week_df = valid[(valid["date"].dt.normalize() >= week_start) & (valid["date"].dt.normalize() <= today)]
    month_df = valid[(valid["date"] >= month_start) & (valid["date"].dt.normalize() <= today)]
    prev_df = valid[(valid["date"] >= prev_month_start) & (valid["date"] < month_start)]
    fy_df = valid[(valid["date"] >= fy_start) & (valid["date"] <= today + pd.Timedelta(days=1))]

    # Calendar days elapsed in the current month; no fabricated future-day average.
    days_elapsed = max(1, today.day)
    return {
        "total_lifetime": total(valid),
        "today": total(today_df),
        "yesterday": total(valid[valid["date"].dt.normalize() == today - pd.Timedelta(days=1)]),
        "this_week": total(week_df),
        "this_month": total(month_df),
        "previous_month": total(prev_df),
        "month_change_pct": ((total(month_df) - total(prev_df)) / abs(total(prev_df)) * 100) if total(prev_df) else None,
        "average_daily_month": total(month_df) / days_elapsed,
        "financial_year": total(fy_df),
        "transaction_count": int(len(valid)),
        "highest_expense": float(valid["amount_inr"].max()) if not valid.empty else 0.0,
        "lowest_expense": float(valid["amount_inr"].min()) if not valid.empty else 0.0,
        "issue_count": 0,
        "valid_transaction_count": int(len(valid)),
    }

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_at TEXT NOT NULL,
                spreadsheet_title TEXT,
                ledger_sheet TEXT,
                transaction_count INTEGER,
                total_expenses REAL,
                issue_count INTEGER,
                metrics_json TEXT,
                status TEXT NOT NULL,
                error TEXT
            )
        """)
        conn.commit()

def save_run(title, sheet, metrics, issue_count, status="SUCCESS", error=None):
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """INSERT INTO runs
            (run_at, spreadsheet_title, ledger_sheet, transaction_count,
             total_expenses, issue_count, metrics_json, status, error)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                datetime.now(ZoneInfo(TIMEZONE)).isoformat(),
                title, sheet, metrics.get("transaction_count", 0),
                metrics.get("total_lifetime", 0), issue_count,
                json.dumps(metrics, default=str), status, error,
            ),
        )
        conn.commit()

def latest_run():
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT run_at, spreadsheet_title, ledger_sheet, transaction_count, "
            "total_expenses, issue_count, status, error FROM runs ORDER BY id DESC LIMIT 1"
        ).fetchone()
    if not row:
        return None
    keys = ["run_at", "spreadsheet_title", "ledger_sheet", "transaction_count",
            "total_expenses", "issue_count", "status", "error"]
    return dict(zip(keys, row))

def run_refresh(record_run=True):
    source_type = os.getenv("SOURCE_TYPE", "excel").strip().lower()
    if source_type == "excel":
        title, raw_sheets = read_workbook()
        source_location = str(get_workbook_path())
    elif source_type == "google_sheets":
        from google_sheets import read_spreadsheet

        title, raw_sheets = read_spreadsheet()
        source_location = "Google Sheets"
    else:
        raise ValueError("SOURCE_TYPE must be 'excel' or 'google_sheets'.")

    ledger_name, raw_ledger = detect_ledger_sheet(raw_sheets)
    df, issues, missing = prepare_transactions(raw_ledger)
    metrics = calculate_metrics(df)
    metrics["issue_count"] = int(len(issues))
    if record_run:
        save_run(title, ledger_name, metrics, len(issues))
    return {
        "title": title,
        "source_type": source_type,
        "source_location": source_location,
        "ledger_sheet": ledger_name,
        "data": df,
        "issues": issues,
        "missing_columns": missing,
        "metrics": metrics,
        "sheet_names": list(raw_sheets.keys()),
        "consolidated_found": any(s.lower().strip() == "consolidated" for s in raw_sheets),
    }

def save_failure(error):
    save_run("", "", {}, 0, status="FAILED", error=str(error))
