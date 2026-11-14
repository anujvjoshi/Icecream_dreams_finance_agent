import os
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from dotenv import load_dotenv

from processor import run_refresh, latest_run

load_dotenv()
TZ = os.getenv("TIMEZONE", "Asia/Kolkata")
SOURCE_TYPE = os.getenv("SOURCE_TYPE", "excel").strip().lower()
WORKBOOK_PATH = os.getenv("WORKBOOK_PATH", "Ice Cream Dreams - Expense Report.xlsx")
REFRESH_SECONDS = int(os.getenv("DASHBOARD_REFRESH_SECONDS", "300"))
auto_refresh_count = st_autorefresh(interval=REFRESH_SECONDS * 1000, key="finance_auto_refresh")

st.set_page_config(
    page_title="Ice Cream Dreams | Expense Intelligence",
    page_icon="🍦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
[data-testid="stMetric"] {background: #f7f9fc; border: 1px solid #e7ebf1;
    padding: 14px 16px; border-radius: 12px;}
.small-muted {color: #64748b; font-size: 0.85rem;}
</style>
""", unsafe_allow_html=True)

def inr(value):
    return f"₹{float(value or 0):,.2f}"

def refresh_data():
    try:
        st.session_state["finance_result"] = run_refresh(record_run=False)
        st.session_state["refresh_error"] = None
        st.session_state["last_loaded"] = datetime.now(ZoneInfo(TZ)).isoformat()
    except Exception as exc:
        st.session_state["refresh_error"] = str(exc)

should_refresh = "finance_result" not in st.session_state
previous_auto_count = st.session_state.get("previous_auto_refresh_count", -1)
if auto_refresh_count != previous_auto_count:
    should_refresh = True
    st.session_state["previous_auto_refresh_count"] = auto_refresh_count

if should_refresh:
    with st.spinner("Opening the configured source and processing the expense ledger..."):
        refresh_data()

with st.sidebar:
    st.title("🍦 Ice Cream Dreams")
    st.caption("Expense Intelligence")
    if st.button("↻ Refresh source data", use_container_width=True):
        with st.spinner("Refreshing source data..."):
            refresh_data()
    st.divider()
    st.caption("Dashboard refresh")
    st.write(f"Every {REFRESH_SECONDS // 60} minute(s), when this page is open")
    st.caption("Scheduled agent run")
    st.write(f"Daily at {os.getenv('SCHEDULE_TIME', '09:00')} {TZ}")
    st.divider()
    st.caption("Configured source")
    st.write(WORKBOOK_PATH if SOURCE_TYPE == "excel" else "Google Sheets")

result = st.session_state.get("finance_result")
error = st.session_state.get("refresh_error")

st.title("Expense Intelligence Dashboard")
st.caption("Live expense monitoring, workbook-wide financial summaries and management insights")

if error:
    st.error(f"Could not refresh source data: {error}")
    previous = latest_run()
    if previous and previous["status"] == "SUCCESS":
        st.warning("The last successful run is recorded in the audit database. Current live data is not available; do not treat this page as freshly reconciled.")
    st.info("Check the configured source path and worksheet. Google Sheets mode also requires OAuth credentials and a spreadsheet ID.")
    st.stop()

if not result:
    st.warning("No data loaded yet. Add Google OAuth credentials and refresh.")
    st.stop()

df = result["data"].copy()
issues = result["issues"].copy()
metrics = result["metrics"]
valid = df[df["is_calculable"]].copy()

loaded = st.session_state.get("last_loaded", "Not recorded")
st.markdown(
    f'<p class="small-muted">Source: {result["title"]} · Ledger tab: {result["ledger_sheet"]} · '
    f'Last successful dashboard refresh: {loaded}</p>',
    unsafe_allow_html=True,
)

# Filters
with st.expander("Filters", expanded=False):
    c1, c2, c3 = st.columns(3)
    min_date = valid["date"].min().date() if not valid.empty else datetime.now().date()
    max_date = valid["date"].max().date() if not valid.empty else datetime.now().date()
    date_range = c1.date_input("Transaction date range", value=(min_date, max_date))
    categories = sorted(valid["category"].dropna().unique().tolist())
    selected_categories = c2.multiselect("Expense categories", categories, default=categories)
    payers = sorted([x for x in valid["paid_by"].dropna().unique().tolist() if x])
    selected_payers = c3.multiselect("Paid by", payers, default=payers)

filtered = valid.copy()
if isinstance(date_range, tuple) and len(date_range) == 2:
    start, end = date_range
    filtered = filtered[(filtered["date"].dt.date >= start) & (filtered["date"].dt.date <= end)]
if selected_categories:
    filtered = filtered[filtered["category"].isin(selected_categories)]
else:
    filtered = filtered.iloc[0:0]
if selected_payers:
    filtered = filtered[filtered["paid_by"].isin(selected_payers) | filtered["paid_by"].eq("")]
elif payers:
    filtered = filtered[filtered["paid_by"].eq("")]

# KPI cards use all ledger data so the headline numbers remain unambiguous.
k1, k2, k3, k4 = st.columns(4)
k1.metric("Total recorded expenses", inr(metrics["total_lifetime"]))
k2.metric("Today's expenses", inr(metrics["today"]))
k3.metric("This month", inr(metrics["this_month"]))
k4.metric("Previous month", inr(metrics["previous_month"]))

k5, k6, k7, k8 = st.columns(4)
change = metrics["month_change_pct"]
k5.metric("Month-on-month change", "N/A" if change is None else f"{change:+.1f}%")
k6.metric("Average daily this month", inr(metrics["average_daily_month"]))
k7.metric("Transactions", f'{metrics["transaction_count"]:,}')
k8.metric("Data quality issues", f'{len(issues):,}')

finance_by_month = result["monthly_finances"]
if not finance_by_month.empty:
    latest_finance = finance_by_month.iloc[-1]
    latest_margin = latest_finance["margin_pct"]
    latest_margin_change = latest_finance["margin_change_pp"]

    st.divider()
    st.subheader("Monthly receipts vs ledger expenses")
    st.caption(
        "Receipts use Kiosk and DNS cash/online fields in DAILY_TALLY; expenses use Monhtly_Expense. "
        "Investments and Prasad account movements are reported separately. Margin is provisional."
    )
    f1, f2, f3, f4, f5 = st.columns(5)
    f1.metric(f"Latest month receipts · {latest_finance['month']}", inr(latest_finance["income_inr"]))
    f2.metric("Ledger expenses", inr(latest_finance["operating_expenses_inr"]))
    f3.metric("Recorded net", inr(latest_finance["surplus_inr"]))
    f4.metric("Provisional margin", "N/A" if pd.isna(latest_margin) else f"{latest_margin:.1f}%")
    f5.metric(
        "MoM margin change",
        "N/A" if pd.isna(latest_margin_change) else f"{latest_margin_change:+.1f} pp",
    )

    finance_left, finance_right = st.columns([1.2, 0.8])
    with finance_left:
        finance_chart = finance_by_month.rename(columns={
            "income_inr": "Reported receipts",
            "operating_expenses_inr": "Recorded ledger expenses",
        })
        fig = px.bar(
            finance_chart,
            x="month",
            y=["Reported receipts", "Recorded ledger expenses"],
            barmode="group",
            labels={"month": "Month", "value": "Amount (INR)", "variable": "Measure"},
        )
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)
    with finance_right:
        margin_chart = finance_by_month.dropna(subset=["margin_pct"])
        if not margin_chart.empty:
            fig = px.line(
                margin_chart,
                x="month",
                y="margin_pct",
                markers=True,
                labels={"month": "Month", "margin_pct": "Provisional margin (%)"},
            )
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Margin is unavailable because receipts are missing for the reported periods.")

    finance_display = finance_by_month[[
        "month", "income_inr", "operating_expenses_inr", "surplus_inr",
        "margin_pct", "margin_change_pp", "kiosk_income_inr", "dns_income_inr",
        "other_reported_income_inr",
    ]].copy()
    finance_display = finance_display.rename(columns={
        "month": "Month",
        "income_inr": "Receipts (INR)",
        "operating_expenses_inr": "Ledger expenses (INR)",
        "surplus_inr": "Surplus (INR)",
        "margin_pct": "Provisional margin (%)",
        "margin_change_pp": "MoM margin change (pp)",
        "kiosk_income_inr": "Kiosk receipts (INR)",
        "dns_income_inr": "DNS receipts (INR)",
        "other_reported_income_inr": "Other reported income (INR, excluded)",
    })
    st.dataframe(finance_display, use_container_width=True, hide_index=True)
    st.caption(
        "Provisional margin = (Kiosk/DNS cash and online receipts − Monhtly_Expense ledger amount) ÷ receipts. "
        "MoM change is shown in percentage points. Other attributed/platform income is excluded "
        "from margin because overlap with receipt totals has not been verified."
    )

st.divider()
left, right = st.columns([1.15, 0.85])

with left:
    st.subheader("Expense by category")
    category = filtered.groupby("category", as_index=False)["amount_inr"].sum().sort_values("amount_inr", ascending=False)
    if not category.empty:
        fig = px.bar(category, x="amount_inr", y="category", orientation="h",
                     labels={"amount_inr": "Amount (INR)", "category": "Category"},
                     text_auto=",.2f")
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=390, margin=dict(l=10,r=10,t=20,b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No transactions match the selected filters.")

with right:
    st.subheader("Category share")
    if not category.empty:
        fig = px.pie(category, names="category", values="amount_inr", hole=0.55)
        fig.update_layout(height=390, margin=dict(l=10,r=10,t=20,b=10), legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No category data available.")

if not filtered.empty:
    category_detail = filtered.groupby("category", as_index=False).agg(
        amount_inr=("amount_inr", "sum"),
        transactions=("amount_inr", "size"),
        average_amount_inr=("amount_inr", "mean"),
    ).sort_values("amount_inr", ascending=False)
    category_total = category_detail["amount_inr"].sum()
    category_detail["share_pct"] = (
        category_detail["amount_inr"] / category_total * 100 if category_total else 0
    )
    category_detail = category_detail.rename(columns={
        "category": "Category",
        "amount_inr": "Spend (INR)",
        "transactions": "Transactions",
        "average_amount_inr": "Average transaction (INR)",
        "share_pct": "Share of filtered spend (%)",
    })
    st.dataframe(category_detail, use_container_width=True, hide_index=True)

st.divider()
left, right = st.columns(2)
with left:
    st.subheader("Monthly expense trend")
    monthly = valid.assign(month=valid["date"].dt.to_period("M").astype(str)).groupby("month", as_index=False)["amount_inr"].sum()
    if not monthly.empty:
        fig = px.line(monthly, x="month", y="amount_inr", markers=True,
                      labels={"month": "Month", "amount_inr": "Amount (INR)"})
        fig.update_layout(height=340, margin=dict(l=10,r=10,t=20,b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No monthly data available.")

with right:
    st.subheader("Paid by")
    payer_rows = filtered.copy()
    payer_rows["paid_by"] = payer_rows["paid_by"].replace("", "Not recorded")
    payer = payer_rows.groupby("paid_by", as_index=False).agg(
        amount_inr=("amount_inr", "sum"),
        transactions=("amount_inr", "size"),
        average_amount_inr=("amount_inr", "mean"),
    ).sort_values("amount_inr", ascending=False)
    payer_total = payer["amount_inr"].sum()
    payer["share_pct"] = payer["amount_inr"] / payer_total * 100 if payer_total else 0
    if not payer.empty:
        fig = px.bar(payer, x="paid_by", y="amount_inr",
                     labels={"paid_by": "Paid by", "amount_inr": "Amount (INR)"})
        fig.update_layout(height=340, margin=dict(l=10,r=10,t=20,b=10))
        st.plotly_chart(fig, use_container_width=True)
        payer = payer.rename(columns={
            "paid_by": "Recorded payer",
            "amount_inr": "Spend (INR)",
            "transactions": "Transactions",
            "average_amount_inr": "Average transaction (INR)",
            "share_pct": "Share of filtered spend (%)",
        })
        st.dataframe(payer, use_container_width=True, hide_index=True)
        st.caption("Paid By is shown as recorded in the ledger; it does not establish funding source or reimbursement status.")
    else:
        st.info("No payer information available.")

st.divider()
st.subheader("Inventory and Snowball spending")
inventory = valid[valid["category"].str.contains("inventory|snowball", case=False, na=False)]
snowball = inventory[inventory["category"].str.contains("snowball", case=False, na=False)]
a, b, c = st.columns(3)
a.metric("Total inventory", inr(inventory["amount_inr"].sum()))
b.metric("Snowball inventory", inr(snowball["amount_inr"].sum()))
c.metric("Inventory transactions", f"{len(inventory):,}")
if not inventory.empty:
    st.dataframe(inventory[["date", "invoice", "category", "particular", "amount_inr", "paid_by", "comments"]],
                 use_container_width=True, hide_index=True)

st.divider()
st.subheader("Transaction ledger")
display = filtered.sort_values("date", ascending=False).copy()
display["date"] = display["date"].dt.strftime("%d-%b-%Y")
display = display.rename(columns={
    "date": "Date", "invoice": "Invoice", "type": "Type", "category": "Normalized category",
    "particular": "Particular", "amount_inr": "Amount (INR)", "paid_by": "Paid By", "comments": "Comments"
})
show_cols = ["Date", "Invoice", "Type", "Normalized category", "Particular", "Amount (INR)", "Paid By", "Comments"]
st.dataframe(display[show_cols], use_container_width=True, hide_index=True)
csv = display[show_cols].to_csv(index=False).encode("utf-8-sig")
st.download_button("Download filtered transactions (CSV)", csv, "ice_cream_dreams_transactions.csv", "text/csv")

st.divider()
st.subheader("Other workbook sheets")
st.dataframe(result["sheet_overview"], use_container_width=True, hide_index=True)
with st.expander("Inspect source worksheet rows", expanded=False):
    selected_sheet = st.selectbox("Worksheet", result["sheet_names"], key="source_sheet_preview")
    source_rows = result["raw_sheets"][selected_sheet]
    header_index = next((
        index for index, row in enumerate(source_rows[:20])
        if sum(value is not None and bool(str(value).strip()) for value in row) >= 3
        and any(str(value).strip().lower() == "date" for value in row if value is not None)
    ), None)
    if header_index is None:
        source_width = max((len(row) for row in source_rows), default=0)
        preview_rows = [
            list(row[:source_width]) + [None] * max(0, source_width - len(row))
            for row in source_rows
        ]
        preview = pd.DataFrame(preview_rows)
        preview.columns = [f"Column {index + 1}" for index in range(len(preview.columns))]
        preview.insert(0, "Source row", range(1, len(preview) + 1))
    else:
        source_width = max((len(row) for row in source_rows), default=0)
        header_row = list(source_rows[header_index])
        header_row.extend([None] * (source_width - len(header_row)))
        headers = []
        for index, value in enumerate(header_row):
            label = str(value).strip() if value is not None and str(value).strip() else f"Column {index + 1}"
            if label in headers:
                label = f"{label} ({index + 1})"
            headers.append(label)
        preview_rows = [
            list(row[:source_width]) + [None] * max(0, source_width - len(row))
            for row in source_rows[header_index + 1:]
        ]
        preview = pd.DataFrame(preview_rows, columns=headers)
        preview.insert(0, "Source row", range(header_index + 2, header_index + 2 + len(preview)))
    preview = preview.dropna(how="all")
    max_preview_rows = st.selectbox("Maximum rows", [100, 500, 1000], index=0, key="source_sheet_preview_limit")
    st.caption("Displays stored/cached worksheet values; formulas are not recalculated by this dashboard.")
    st.dataframe(preview.head(max_preview_rows), use_container_width=True, hide_index=True)

supplemental_spending = result["supplemental_spending"]
if not supplemental_spending.empty:
    st.subheader("Investment and vendor spending")
    st.caption("Kept separate from operating expenses; vendor sheets may overlap other records.")
    investment_summary = supplemental_spending.copy()
    investment_summary["paid_by"] = investment_summary["paid_by"].replace("", "Not recorded")
    investment_summary = investment_summary.groupby(
        ["source_sheet", "paid_by"], as_index=False
    ).agg(
        amount_inr=("amount_inr", "sum"),
        transactions=("amount_inr", "size"),
    ).sort_values("amount_inr", ascending=False)
    total_supplemental_spend = investment_summary["amount_inr"].sum()
    investment_summary["share_pct"] = (
        investment_summary["amount_inr"] / total_supplemental_spend * 100
        if total_supplemental_spend else 0
    )
    investment_summary = investment_summary.rename(columns={
        "source_sheet": "Worksheet",
        "paid_by": "Recorded payer",
        "amount_inr": "Recorded amount (INR)",
        "transactions": "Transactions",
        "share_pct": "Share of investment/vendor spend (%)",
    })
    st.dataframe(investment_summary, use_container_width=True, hide_index=True)

prasad_activity = result["prasad_activity"].copy()
prasad_activity = prasad_activity[
    prasad_activity["date"].notna()
    & (prasad_activity["received_inr"].notna() | prasad_activity["spent_inr"].notna())
]
if not prasad_activity.empty:
    st.subheader("Prasad account movements")
    st.caption("Reported receipts and spending in this account ledger; excluded from sales and operating margin.")
    prasad_activity["month"] = prasad_activity["date"].dt.to_period("M").astype(str)
    prasad_monthly = prasad_activity.groupby("month", as_index=False).agg(
        received_inr=("received_inr", "sum"),
        spent_inr=("spent_inr", "sum"),
    ).rename(columns={
        "month": "Month",
        "received_inr": "Received (INR)",
        "spent_inr": "Spent (INR)",
    })
    st.dataframe(prasad_monthly, use_container_width=True, hide_index=True)

st.divider()
st.subheader("Data quality and reconciliation")
if issues.empty:
    st.success("No validation issues detected in the current ledger scan.")
else:
    st.warning(f"{len(issues)} issue(s) need review. Original source rows have not been changed.")
    st.dataframe(issues, use_container_width=True, hide_index=True)

if result["missing_columns"]:
    st.info("Optional/unmapped source columns: " + ", ".join(result["missing_columns"]))

if result["consolidated_found"]:
    st.caption("A Consolidated tab was found. This initial build detects its presence; a verified total-to-total reconciliation requires confirming its summary layout and formulas.")
else:
    st.caption("No tab named 'Consolidated' was detected. The dashboard calculates directly from the detected transaction ledger.")

st.divider()
st.subheader("Management summary")
top_category = category.iloc[0] if not category.empty else None
if top_category is not None:
    st.write(
        f"The largest category in the selected date/filter view is **{top_category['category']}** "
        f"at **{inr(top_category['amount_inr'])}**."
    )
else:
    st.write("No transactions are available for the selected view.")
if change is not None:
    direction = "increased" if change > 0 else "decreased"
    st.write(f"Recorded expenses this month have {direction} by **{abs(change):.1f}%** compared with the previous calendar month.")
else:
    st.write("Month-on-month comparison is unavailable because the previous month has no recorded expenses.")

st.caption("All totals are based on recorded ledger amounts. This workbook does not establish revenue, profit, unpaid liabilities or cash balance.")
