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
st.caption("Live expense monitoring, reconciliation and management insights")

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
    payer = filtered[filtered["paid_by"].ne("")].groupby("paid_by", as_index=False)["amount_inr"].sum().sort_values("amount_inr", ascending=False)
    if not payer.empty:
        fig = px.bar(payer, x="paid_by", y="amount_inr",
                     labels={"paid_by": "Paid by", "amount_inr": "Amount (INR)"})
        fig.update_layout(height=340, margin=dict(l=10,r=10,t=20,b=10))
        st.plotly_chart(fig, use_container_width=True)
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
