import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
from datetime import timedelta


# =========================
# PAGE CONFIG
# =========================

st.set_page_config(
    page_title="Inventory Activity Dashboard",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Inventory Activity Dashboard")
st.caption(
    "Upload Item Activity Report Excel file. "
    "Dashboard will automatically clean data, calculate inventory movement, shortage risk, and forecast."
)


# =========================
# HELPER FUNCTIONS
# =========================

def extract_number(value):
    """
    Extract first number from values like:
    35.0000 / 35
    1.0000 / 1
    blank

    This returns the unit quantity.
    Example:
    5139 / 198 -> 5139
    """
    if pd.isna(value):
        return 0.0

    text = str(value).strip()

    if text == "":
        return 0.0

    match = re.search(r"[-+]?\d*\.?\d+", text)

    if match:
        return float(match.group())

    return 0.0


def extract_ctn_number(value):
    """
    Extract carton/CTN number from values like:
    5139 / 198 -> 198

    If there is no slash value, returns 0.
    """
    if pd.isna(value):
        return 0.0

    text = str(value).strip()

    if "/" not in text:
        return 0.0

    parts = text.split("/")

    if len(parts) < 2:
        return 0.0

    ctn_text = parts[1].strip()

    match = re.search(r"[-+]?\d*\.?\d+", ctn_text)

    if match:
        return float(match.group())

    return 0.0


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def safe_row_text(row):
    """
    Convert every cell in a row to safe lowercase text.
    """
    row_values = []

    for value in row.tolist():
        if pd.isna(value):
            row_values.append("")
        else:
            row_values.append(str(value).strip().lower())

    return " ".join(row_values)


def parse_activity_date(activity_text):
    """
    Parse activity date safely.

    Handles:
    6/1/2026
    6/1/2026 (Not Shipped)
    2026-06-01
    """
    activity_text = clean_text(activity_text)

    if activity_text == "":
        return pd.NaT

    date_match = re.search(
        r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}",
        activity_text
    )

    if date_match:
        return pd.to_datetime(date_match.group(0), errors="coerce")

    return pd.to_datetime(activity_text, errors="coerce")


def find_header_row(raw_df):
    """
    Find the transaction header row automatically.
    """
    for i in range(min(80, len(raw_df))):
        row_text = safe_row_text(raw_df.iloc[i])

        if "sku" in row_text and "activity date" in row_text:
            return i

    return None


def extract_report_range(raw_df):
    """
    Extract report date range from the top header area.

    Example:
    Item Activity From: 2025-09-01 to 2026-06-01
    """
    report_start = None
    report_end = None

    for i in range(min(30, len(raw_df))):
        row = raw_df.iloc[i]
        row_text = safe_row_text(row)

        if "item activity from" in row_text:
            dates_found = []

            for value in row.tolist():
                if pd.isna(value):
                    continue

                parsed_date = pd.to_datetime(value, errors="coerce")

                if pd.notna(parsed_date):
                    dates_found.append(parsed_date)

            if len(dates_found) < 2:
                date_patterns = re.findall(
                    r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}",
                    row_text
                )

                for date_text in date_patterns:
                    parsed_date = pd.to_datetime(date_text, errors="coerce")

                    if pd.notna(parsed_date):
                        dates_found.append(parsed_date)

            if len(dates_found) >= 1:
                report_start = dates_found[0]

            if len(dates_found) >= 2:
                report_end = dates_found[1]

            break

    return report_start, report_end


def set_risk_filter(values):
    """
    Keep KPI buttons and shortage tab filter aligned.
    """
    st.session_state["risk_filter_click"] = values
    st.session_state["shortage_risk_filter"] = values


def filter_recent_window(tx, end_date, days):
    """
    True calendar day window.

    Last 30 days:
    start = end_date - 29 days
    because end date is included.

    Example:
    End date = 2026-06-01
    Last 30 days = 2026-05-03 to 2026-06-01
    """
    start_date = end_date - timedelta(days=days - 1)

    return tx[
        (tx["Activity Date"] >= start_date)
        & (tx["Activity Date"] <= end_date)
    ].copy(), start_date, end_date


def parse_item_activity_report(uploaded_file):
    """
    Main parser for Item Activity Report.

    Expected structure:
    - SKU appears once per item block
    - Transaction rows below SKU have blank SKU
    - Activity Date / Trans# / Ref# / Qty In / Qty Out / Balance
    - Ending Balance row
    - Total row where Ref # = Total
    """

    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    report_start, report_end = extract_report_range(raw)

    header_row = find_header_row(raw)

    if header_row is None:
        raise ValueError(
            "Cannot find the header row. Please make sure this is an Item Activity Report."
        )

    records = []

    current_sku = None
    current_description = None
    current_packed = None

    for idx in range(header_row + 1, len(raw)):
        row = raw.iloc[idx]

        sku_cell = row.iloc[0] if len(row) > 0 else None
        description_cell = row.iloc[2] if len(row) > 2 else None
        packed_cell = row.iloc[6] if len(row) > 6 else None

        # Detect a new SKU block
        if pd.notna(sku_cell) and str(sku_cell).strip() != "":
            sku_text = str(sku_cell).strip()

            ignored_words = [
                "sku",
                "nan",
                "warehouse",
                "customer",
                "item activity",
                "activity date"
            ]

            if sku_text.lower() not in ignored_words:
                current_sku = sku_text
                current_description = clean_text(description_cell)
                current_packed = extract_number(packed_cell)

            continue

        # Transaction columns based on current file layout
        activity_date_raw = row.iloc[7] if len(row) > 7 else None
        trans_no = row.iloc[9] if len(row) > 9 else None
        ref_no = row.iloc[10] if len(row) > 10 else None
        qty_in = row.iloc[12] if len(row) > 12 else None
        qty_out = row.iloc[14] if len(row) > 14 else None
        balance = row.iloc[19] if len(row) > 19 else None
        ctn_balance = row.iloc[20] if len(row) > 20 else None

        activity_text = clean_text(activity_date_raw)
        trans_text = clean_text(trans_no)
        ref_text = clean_text(ref_no)

        activity_text_lower = activity_text.lower()
        ref_text_lower = ref_text.lower()

        is_beginning_balance = activity_text_lower == "beginning balance"
        is_ending_balance = activity_text_lower == "ending balance"

        # In this report, official Total row is in Ref # column.
        is_total_row = ref_text_lower == "total"

        should_keep_row = (
            current_sku is not None
            and (
                activity_text != ""
                or is_total_row
            )
        )

        if not should_keep_row:
            continue

        if is_beginning_balance or is_ending_balance or is_total_row:
            activity_date = pd.NaT
        else:
            activity_date = parse_activity_date(activity_text)

        qty_in_num = extract_number(qty_in)
        qty_out_num = extract_number(qty_out)

        qty_in_ctn = extract_ctn_number(qty_in)
        qty_out_ctn = extract_ctn_number(qty_out)

        balance_num = extract_number(balance)
        ctn_balance_num = extract_number(ctn_balance)

        if is_beginning_balance:
            movement_type = "Beginning Balance"
        elif is_ending_balance:
            movement_type = "Ending Balance"
        elif is_total_row:
            movement_type = "Total"
        elif qty_in_num > 0 and qty_out_num == 0:
            movement_type = "Inbound"
        elif qty_out_num > 0 and qty_in_num == 0:
            movement_type = "Outbound"
        elif qty_in_num > 0 and qty_out_num > 0:
            movement_type = "Mixed"
        else:
            movement_type = "No Movement"

        is_cancelled = (
            "cancel" in ref_text.lower()
            or "cancel" in trans_text.lower()
        )

        is_not_shipped = "not shipped" in activity_text.lower()

        records.append({
            "Source Row": idx + 1,
            "SKU": current_sku,
            "Description": current_description,
            "Packed": current_packed,
            "Activity Date": activity_date,
            "Activity Text": activity_text,
            "Trans #": trans_text,
            "Ref #": ref_text,
            "Qty In": qty_in_num,
            "Qty In CTN": qty_in_ctn,
            "Qty Out": qty_out_num,
            "Qty Out CTN": qty_out_ctn,
            "Balance": balance_num,
            "Ctn Balance": ctn_balance_num,
            "Movement Type": movement_type,
            "Is Beginning Balance": is_beginning_balance,
            "Is Ending Balance": is_ending_balance,
            "Is Total Row": is_total_row,
            "Is Cancelled": is_cancelled,
            "Is Not Shipped": is_not_shipped
        })

    df = pd.DataFrame(records)

    if df.empty:
        raise ValueError("No activity records found.")

    return df, report_start, report_end


def build_summary(df, report_start=None, report_end=None):
    """
    Build SKU-level summary.

    Correct logic:
    - Ending Balance comes from official Ending Balance row
    - Total Inbound / Outbound comes from official Total row where Ref # = Total
    - Recent usage comes from actual dated transaction rows
    - Forecast uses Ending Balance + recent usage
    """

    tx = df[
        (df["Activity Date"].notna())
        & (~df["Movement Type"].isin(["Beginning Balance", "Ending Balance", "Total"]))
    ].copy()

    if tx.empty:
        raise ValueError("No valid transaction dates found.")

    activity_min_date = tx["Activity Date"].min()
    activity_max_date = tx["Activity Date"].max()

    report_min_date = (
        report_start
        if report_start is not None and pd.notna(report_start)
        else activity_min_date
    )

    report_max_date = (
        report_end
        if report_end is not None and pd.notna(report_end)
        else activity_max_date
    )

    # Use report end date as the recent outbound reference date.
    recent_end_date = report_max_date

    # =========================
    # OFFICIAL ENDING BALANCE
    # =========================

    ending_balance_rows = df[df["Movement Type"] == "Ending Balance"].copy()

    if not ending_balance_rows.empty:
        ending_balance_rows["Source Order"] = range(len(ending_balance_rows))

        latest_balance = (
            ending_balance_rows
            .sort_values(["SKU", "Source Order"])
            .groupby("SKU", as_index=False)
            .tail(1)[["SKU", "Description", "Packed", "Balance", "Ctn Balance"]]
        )
    else:
        sorted_df = tx.copy()
        sorted_df["Source Order"] = range(len(sorted_df))

        latest_balance = (
            sorted_df.sort_values(["SKU", "Activity Date", "Source Order"])
            .groupby("SKU", as_index=False)
            .tail(1)[["SKU", "Description", "Packed", "Balance", "Ctn Balance"]]
        )

    latest_balance = latest_balance.rename(columns={
        "Balance": "Ending Balance",
        "Ctn Balance": "Ending Ctn Balance"
    })

    # =========================
    # OFFICIAL TOTAL ROW
    # =========================

    total_rows = df[df["Movement Type"] == "Total"].copy()

    if not total_rows.empty:
        total_rows["Source Order"] = range(len(total_rows))

        official_totals = (
            total_rows
            .sort_values(["SKU", "Source Order"])
            .groupby("SKU", as_index=False)
            .tail(1)[[
                "SKU",
                "Qty In",
                "Qty In CTN",
                "Qty Out",
                "Qty Out CTN",
                "Source Row"
            ]]
            .rename(columns={
                "Qty In": "Total Inbound",
                "Qty In CTN": "Total Inbound CTN",
                "Qty Out": "Total Outbound",
                "Qty Out CTN": "Total Outbound CTN",
                "Source Row": "Official Total Source Row"
            })
        )

        total_inbound = official_totals[["SKU", "Total Inbound", "Total Inbound CTN"]]
        total_outbound = official_totals[["SKU", "Total Outbound", "Total Outbound CTN"]]
        total_source_rows = official_totals[["SKU", "Official Total Source Row"]]

    else:
        total_inbound = (
            tx.groupby("SKU", as_index=False)[["Qty In", "Qty In CTN"]]
            .sum()
            .rename(columns={
                "Qty In": "Total Inbound",
                "Qty In CTN": "Total Inbound CTN"
            })
        )

        total_outbound = (
            tx.groupby("SKU", as_index=False)[["Qty Out", "Qty Out CTN"]]
            .sum()
            .rename(columns={
                "Qty Out": "Total Outbound",
                "Qty Out CTN": "Total Outbound CTN"
            })
        )

        total_source_rows = pd.DataFrame({
            "SKU": total_outbound["SKU"],
            "Official Total Source Row": ""
        })

    # =========================
    # LAST ACTIVITY DATE
    # =========================

    last_activity = (
        tx.groupby("SKU", as_index=False)["Activity Date"]
        .max()
        .rename(columns={"Activity Date": "Last Activity Date"})
    )

    # =========================
    # RECENT OUTBOUND WINDOWS
    # =========================

    tx_30, outbound_30_start, outbound_30_end = filter_recent_window(
        tx,
        recent_end_date,
        30
    )

    tx_14, outbound_14_start, outbound_14_end = filter_recent_window(
        tx,
        recent_end_date,
        14
    )

    tx_7, outbound_7_start, outbound_7_end = filter_recent_window(
        tx,
        recent_end_date,
        7
    )

    outbound_30 = (
        tx_30.groupby("SKU", as_index=False)[["Qty Out", "Qty Out CTN"]]
        .sum()
        .rename(columns={
            "Qty Out": "Outbound Last 30 Days",
            "Qty Out CTN": "Outbound Last 30 Days CTN"
        })
    )

    outbound_14 = (
        tx_14.groupby("SKU", as_index=False)[["Qty Out", "Qty Out CTN"]]
        .sum()
        .rename(columns={
            "Qty Out": "Outbound Last 14 Days",
            "Qty Out CTN": "Outbound Last 14 Days CTN"
        })
    )

    outbound_7 = (
        tx_7.groupby("SKU", as_index=False)[["Qty Out", "Qty Out CTN"]]
        .sum()
        .rename(columns={
            "Qty Out": "Outbound Last 7 Days",
            "Qty Out CTN": "Outbound Last 7 Days CTN"
        })
    )

    # =========================
    # MERGE SUMMARY
    # =========================

    summary = latest_balance.copy()

    for small_df in [
        total_inbound,
        total_outbound,
        total_source_rows,
        last_activity,
        outbound_30,
        outbound_14,
        outbound_7
    ]:
        summary = summary.merge(small_df, on="SKU", how="left")

    fill_cols = [
        "Total Inbound",
        "Total Inbound CTN",
        "Total Outbound",
        "Total Outbound CTN",
        "Outbound Last 30 Days",
        "Outbound Last 30 Days CTN",
        "Outbound Last 14 Days",
        "Outbound Last 14 Days CTN",
        "Outbound Last 7 Days",
        "Outbound Last 7 Days CTN"
    ]

    summary[fill_cols] = summary[fill_cols].fillna(0)

    summary["Avg Daily Usage 30D"] = summary["Outbound Last 30 Days"] / 30

    summary["Days Remaining"] = np.where(
        summary["Avg Daily Usage 30D"] > 0,
        summary["Ending Balance"] / summary["Avg Daily Usage 30D"],
        np.nan
    )

    def risk_level(row):
        ending_balance = row["Ending Balance"]
        usage = row["Avg Daily Usage 30D"]
        days = row["Days Remaining"]

        if ending_balance <= 0 and usage > 0:
            return "Critical"
        if usage == 0:
            return "No Recent Movement"
        if days <= 7:
            return "Critical"
        if days <= 14:
            return "Warning"
        if days <= 30:
            return "Watch"
        return "Safe"

    summary["Risk Level"] = summary.apply(risk_level, axis=1)

    def stockout_date(row):
        if pd.isna(row["Days Remaining"]):
            return ""
        if row["Days Remaining"] < 0:
            return ""

        return (recent_end_date + timedelta(days=float(row["Days Remaining"]))).date()

    summary["Forecast Stockout Date"] = summary.apply(stockout_date, axis=1)

    summary["Report Start Date"] = report_min_date.date()
    summary["Report End Date"] = report_max_date.date()
    summary["Activity Start Date"] = activity_min_date.date()
    summary["Activity End Date"] = activity_max_date.date()

    summary["Outbound 30D Start"] = outbound_30_start.date()
    summary["Outbound 30D End"] = outbound_30_end.date()
    summary["Outbound 14D Start"] = outbound_14_start.date()
    summary["Outbound 14D End"] = outbound_14_end.date()
    summary["Outbound 7D Start"] = outbound_7_start.date()
    summary["Outbound 7D End"] = outbound_7_end.date()

    recent_windows = {
        "30D": {
            "start": outbound_30_start,
            "end": outbound_30_end,
            "rows": tx_30
        },
        "14D": {
            "start": outbound_14_start,
            "end": outbound_14_end,
            "rows": tx_14
        },
        "7D": {
            "start": outbound_7_start,
            "end": outbound_7_end,
            "rows": tx_7
        }
    }

    return (
        summary,
        tx,
        recent_windows,
        report_min_date,
        report_max_date,
        activity_min_date,
        activity_max_date
    )


def apply_risk_sort(df):
    risk_order = {
        "Critical": 1,
        "Warning": 2,
        "Watch": 3,
        "Safe": 4,
        "No Recent Movement": 5
    }

    df = df.copy()
    df["Risk Sort"] = df["Risk Level"].map(risk_order)

    return df.sort_values(
        ["Risk Sort", "Days Remaining"],
        ascending=[True, True]
    )


# =========================
# SIDEBAR
# =========================

st.sidebar.header("How to use")
st.sidebar.write(
    """
    1. Upload Item Activity Report Excel  
    2. Review KPI cards  
    3. Click risk buttons to view SKUs  
    4. Check Shortage Forecast tab  
    5. Search SKU if needed  
    6. Download processed CSV  
    """
)

st.sidebar.divider()

st.sidebar.write("### Balance / Total Logic")
st.sidebar.write(
    """
    Ending Balance = official Ending Balance row  
    Ending Ctn Balance = official Ending Balance row  
    Total Inbound = official Total row where Ref # = Total  
    Total Outbound = official Total row where Ref # = Total  
    Recent usage = dated transaction rows  
    """
)

st.sidebar.divider()

st.sidebar.write("### Recent Outbound Logic")
st.sidebar.write(
    """
    Last 30 days = Report End Date - 29 days through Report End Date  
    Last 14 days = Report End Date - 13 days through Report End Date  
    Last 7 days = Report End Date - 6 days through Report End Date  
    Rows like 6/1/2026 (Not Shipped) are included if Qty Out exists.  
    """
)

st.sidebar.divider()

st.sidebar.write("### Risk Level Logic")
st.sidebar.write(
    """
    Critical: 0-7 days remaining  
    Warning: 8-14 days remaining  
    Watch: 15-30 days remaining  
    Safe: More than 30 days remaining  
    No Recent Movement: No outbound in last 30 days  
    """
)


uploaded_file = st.file_uploader(
    "Upload / Drop Excel file here",
    type=["xlsx"],
    help="Upload Item Activity Report exported from WMS."
)


# =========================
# MAIN APP
# =========================

if uploaded_file is None:
    st.info("Please upload an Item Activity Report Excel file to start.")

else:
    try:
        df, report_start, report_end = parse_item_activity_report(uploaded_file)

        (
            summary,
            tx,
            recent_windows,
            report_min_date,
            report_max_date,
            activity_min_date,
            activity_max_date
        ) = build_summary(df, report_start, report_end)

        st.success(
            f"File processed successfully. "
            f"Report range: {report_min_date.date()} to {report_max_date.date()}"
        )

        with st.expander("Date Range Details"):
            st.write("**Report range from file header:**")
            st.write(f"{report_min_date.date()} to {report_max_date.date()}")

            st.write("**Actual activity date range found in transactions:**")
            st.write(f"{activity_min_date.date()} to {activity_max_date.date()}")

            st.write("**Recent outbound windows used:**")
            st.write(
                f"30D: {recent_windows['30D']['start'].date()} to {recent_windows['30D']['end'].date()}"
            )
            st.write(
                f"14D: {recent_windows['14D']['start'].date()} to {recent_windows['14D']['end'].date()}"
            )
            st.write(
                f"7D: {recent_windows['7D']['start'].date()} to {recent_windows['7D']['end'].date()}"
            )

            st.caption(
                "Recent outbound uses report end date and true calendar day windows."
            )

        with st.expander("Balance, Total, and Recent Outbound Logic Details"):
            st.write("**Ending Balance source:**")
            st.write("The dashboard uses the official `Ending Balance` row for each SKU.")

            st.write("**Total Inbound / Total Outbound source:**")
            st.write("The dashboard uses the official `Total` row where `Ref # = Total`.")

            st.write("**Recent usage source:**")
            st.write(
                "Outbound Last 7/14/30 Days is calculated from actual dated transaction rows, "
                "including rows like `6/1/2026 (Not Shipped)` if they have Qty Out."
            )

            st.write("**Forecast calculation:**")
            st.write("Shortage risk and Days Remaining are calculated using `Ending Balance` and recent outbound usage.")

        # =========================
        # SESSION STATE
        # =========================

        if "risk_filter_click" not in st.session_state:
            st.session_state["risk_filter_click"] = [
                "Critical",
                "Warning",
                "Watch"
            ]

        # =========================
        # KPI CARDS
        # =========================

        total_skus = summary["SKU"].nunique()
        total_ending_balance = summary["Ending Balance"].sum()
        total_ending_ctn_balance = summary["Ending Ctn Balance"].sum()
        total_inbound = summary["Total Inbound"].sum()
        total_inbound_ctn = summary["Total Inbound CTN"].sum()
        total_outbound = summary["Total Outbound"].sum()
        total_outbound_ctn = summary["Total Outbound CTN"].sum()
        total_outbound_30 = summary["Outbound Last 30 Days"].sum()
        total_outbound_14 = summary["Outbound Last 14 Days"].sum()
        total_outbound_7 = summary["Outbound Last 7 Days"].sum()
        critical_count = (summary["Risk Level"] == "Critical").sum()
        warning_count = (summary["Risk Level"] == "Warning").sum()
        watch_count = (summary["Risk Level"] == "Watch").sum()
        safe_count = (summary["Risk Level"] == "Safe").sum()
        no_movement_count = (summary["Risk Level"] == "No Recent Movement").sum()

        k1, k2, k3, k4, k5, k6 = st.columns(6)

        with k1:
            st.metric("Total SKUs", f"{total_skus:,}")
            if st.button("View All SKUs", use_container_width=True):
                set_risk_filter([
                    "Critical",
                    "Warning",
                    "Watch",
                    "Safe",
                    "No Recent Movement"
                ])

        with k2:
            st.metric("Ending Balance", f"{total_ending_balance:,.0f}")
            if st.button("View Safe SKUs", use_container_width=True):
                set_risk_filter(["Safe"])

        with k3:
            st.metric("Total Inbound", f"{total_inbound:,.0f} / {total_inbound_ctn:,.0f} CTN")
            if st.button("View No Movement", use_container_width=True):
                set_risk_filter(["No Recent Movement"])

        with k4:
            st.metric("Total Outbound", f"{total_outbound:,.0f} / {total_outbound_ctn:,.0f} CTN")
            if st.button("View Watch SKUs", use_container_width=True):
                set_risk_filter(["Watch"])

        with k5:
            st.metric("Critical SKUs", f"{critical_count:,}")
            if st.button("View Critical SKUs", use_container_width=True):
                set_risk_filter(["Critical"])

        with k6:
            st.metric("Warning SKUs", f"{warning_count:,}")
            if st.button("View Warning SKUs", use_container_width=True):
                set_risk_filter(["Warning"])

        # =========================
        # RECENT OUTBOUND KPI
        # =========================

        st.subheader("Recent Outbound Summary")

        r1, r2, r3 = st.columns(3)

        with r1:
            st.metric(
                "Outbound Last 30 Days",
                f"{total_outbound_30:,.0f}",
                help=f"{recent_windows['30D']['start'].date()} to {recent_windows['30D']['end'].date()}"
            )

        with r2:
            st.metric(
                "Outbound Last 14 Days",
                f"{total_outbound_14:,.0f}",
                help=f"{recent_windows['14D']['start'].date()} to {recent_windows['14D']['end'].date()}"
            )

        with r3:
            st.metric(
                "Outbound Last 7 Days",
                f"{total_outbound_7:,.0f}",
                help=f"{recent_windows['7D']['start'].date()} to {recent_windows['7D']['end'].date()}"
            )

        # =========================
        # INTERACTIVE QUICK VIEW
        # =========================

        st.divider()

        quick_view = summary[
            summary["Risk Level"].isin(st.session_state["risk_filter_click"])
        ].copy()

        quick_view = apply_risk_sort(quick_view)

        quick_cols = [
            "SKU",
            "Description",
            "Packed",
            "Ending Balance",
            "Ending Ctn Balance",
            "Total Inbound",
            "Total Inbound CTN",
            "Total Outbound",
            "Total Outbound CTN",
            "Official Total Source Row",
            "Outbound Last 30 Days",
            "Outbound Last 30 Days CTN",
            "Outbound Last 14 Days",
            "Outbound Last 14 Days CTN",
            "Outbound Last 7 Days",
            "Outbound Last 7 Days CTN",
            "Avg Daily Usage 30D",
            "Days Remaining",
            "Forecast Stockout Date",
            "Risk Level",
            "Last Activity Date"
        ]

        selected_filter_label = ", ".join(st.session_state["risk_filter_click"])

        st.subheader(f"Quick SKU View: {selected_filter_label}")

        q1, q2, q3, q4, q5 = st.columns(5)

        q1.metric("Critical", f"{critical_count:,}")
        q2.metric("Warning", f"{warning_count:,}")
        q3.metric("Watch", f"{watch_count:,}")
        q4.metric("Safe", f"{safe_count:,}")
        q5.metric("No Movement", f"{no_movement_count:,}")

        st.dataframe(
            quick_view[quick_cols],
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "Download Current Quick View CSV",
            quick_view[quick_cols].to_csv(index=False).encode("utf-8"),
            file_name="quick_sku_view.csv",
            mime="text/csv"
        )

        # =========================
        # TABS
        # =========================

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "Overview",
            "Shortage Forecast",
            "Movement Trend",
            "SKU Detail",
            "Exceptions",
            "Recent Outbound Audit"
        ])

        # =========================
        # TAB 1: OVERVIEW
        # =========================

        with tab1:
            st.subheader("Inventory Overview")

            c1, c2 = st.columns(2)

            with c1:
                risk_counts = (
                    summary["Risk Level"]
                    .value_counts()
                    .reset_index()
                )

                risk_counts.columns = ["Risk Level", "SKU Count"]

                fig_risk = px.pie(
                    risk_counts,
                    names="Risk Level",
                    values="SKU Count",
                    title="SKU Risk Distribution"
                )

                st.plotly_chart(fig_risk, use_container_width=True)

            with c2:
                top_outbound = summary.sort_values(
                    "Total Outbound",
                    ascending=False
                ).head(15)

                fig_top = px.bar(
                    top_outbound,
                    x="SKU",
                    y="Total Outbound",
                    hover_data=[
                        "Description",
                        "Ending Balance",
                        "Ending Ctn Balance",
                        "Official Total Source Row",
                        "Risk Level"
                    ],
                    title="Top 15 Official Total Outbound SKUs"
                )

                st.plotly_chart(fig_top, use_container_width=True)

            st.subheader("SKU Summary Table")

            st.dataframe(
                summary.sort_values("Total Outbound", ascending=False),
                use_container_width=True,
                hide_index=True
            )

        # =========================
        # TAB 2: SHORTAGE FORECAST
        # =========================

        with tab2:
            st.subheader("Shortage Forecast")

            risk_options = [
                "Critical",
                "Warning",
                "Watch",
                "Safe",
                "No Recent Movement"
            ]

            selected_risks = st.multiselect(
                "Risk Level Filter",
                risk_options,
                default=st.session_state["risk_filter_click"],
                key="shortage_risk_filter"
            )

            shortage_view = summary[
                summary["Risk Level"].isin(selected_risks)
            ].copy()

            shortage_view = apply_risk_sort(shortage_view)

            cols_to_show = [
                "SKU",
                "Description",
                "Packed",
                "Ending Balance",
                "Ending Ctn Balance",
                "Total Inbound",
                "Total Inbound CTN",
                "Total Outbound",
                "Total Outbound CTN",
                "Official Total Source Row",
                "Outbound Last 30 Days",
                "Outbound Last 30 Days CTN",
                "Outbound Last 14 Days",
                "Outbound Last 14 Days CTN",
                "Outbound Last 7 Days",
                "Outbound Last 7 Days CTN",
                "Avg Daily Usage 30D",
                "Days Remaining",
                "Forecast Stockout Date",
                "Risk Level",
                "Last Activity Date",
                "Report Start Date",
                "Report End Date",
                "Activity Start Date",
                "Activity End Date",
                "Outbound 30D Start",
                "Outbound 30D End",
                "Outbound 14D Start",
                "Outbound 14D End",
                "Outbound 7D Start",
                "Outbound 7D End"
            ]

            st.dataframe(
                shortage_view[cols_to_show],
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                "Download Shortage Forecast CSV",
                shortage_view[cols_to_show].to_csv(index=False).encode("utf-8"),
                file_name="shortage_forecast.csv",
                mime="text/csv"
            )

        # =========================
        # TAB 3: MOVEMENT TREND
        # =========================

        with tab3:
            st.subheader("Inventory Movement Trend")

            daily = (
                tx.groupby("Activity Date", as_index=False)[["Qty In", "Qty Out"]]
                .sum()
                .sort_values("Activity Date")
            )

            daily["Net Movement"] = daily["Qty In"] - daily["Qty Out"]

            daily_melt = daily.melt(
                id_vars="Activity Date",
                value_vars=["Qty In", "Qty Out"],
                var_name="Type",
                value_name="Quantity"
            )

            fig_daily = px.line(
                daily_melt,
                x="Activity Date",
                y="Quantity",
                color="Type",
                markers=True,
                title="Daily Inbound vs Outbound"
            )

            st.plotly_chart(fig_daily, use_container_width=True)

            fig_net = px.bar(
                daily,
                x="Activity Date",
                y="Net Movement",
                title="Daily Net Movement"
            )

            st.plotly_chart(fig_net, use_container_width=True)

            st.dataframe(
                daily,
                use_container_width=True,
                hide_index=True
            )

        # =========================
        # TAB 4: SKU DETAIL
        # =========================

        with tab4:
            st.subheader("SKU Detail Lookup")

            sku_list = sorted(summary["SKU"].dropna().unique())

            selected_sku = st.selectbox(
                "Select SKU",
                sku_list
            )

            sku_summary = summary[summary["SKU"] == selected_sku]

            if not sku_summary.empty:
                st.write("### SKU Summary")

                st.dataframe(
                    sku_summary,
                    use_container_width=True,
                    hide_index=True
                )

            sku_tx = tx[tx["SKU"] == selected_sku].sort_values("Activity Date")

            st.write("### Transaction History")

            st.dataframe(
                sku_tx,
                use_container_width=True,
                hide_index=True
            )

            if not sku_tx.empty:
                fig_sku_balance = px.line(
                    sku_tx,
                    x="Activity Date",
                    y="Balance",
                    markers=True,
                    hover_data=[
                        "Trans #",
                        "Ref #",
                        "Qty In",
                        "Qty Out",
                        "Is Not Shipped"
                    ],
                    title=f"Transaction Balance Trend - {selected_sku}"
                )

                st.plotly_chart(fig_sku_balance, use_container_width=True)

        # =========================
        # TAB 5: EXCEPTIONS
        # =========================

        with tab5:
            st.subheader("Exception / Audit Report")

            st.write("### Cancelled Transactions")

            cancelled = df[df["Is Cancelled"] == True]

            st.dataframe(
                cancelled,
                use_container_width=True,
                hide_index=True
            )

            st.write("### Not Shipped Rows with Qty Out")

            not_shipped = df[
                (df["Is Not Shipped"] == True)
                & (df["Qty Out"] > 0)
            ]

            st.dataframe(
                not_shipped,
                use_container_width=True,
                hide_index=True
            )

            st.write("### Negative Ending Ctn Balance")

            negative_ctn = summary[
                summary["Ending Ctn Balance"] < 0
            ]

            st.dataframe(
                negative_ctn,
                use_container_width=True,
                hide_index=True
            )

            st.write("### Zero Ending Balance with Recent Outbound")

            zero_recent = summary[
                (summary["Ending Balance"] <= 0)
                & (summary["Outbound Last 30 Days"] > 0)
            ]

            st.dataframe(
                zero_recent,
                use_container_width=True,
                hide_index=True
            )

            st.write("### Fast Moving but Low Ending Balance")

            fast_low = summary[
                (summary["Outbound Last 30 Days"] > 0)
                & (summary["Days Remaining"] <= 30)
            ].sort_values("Days Remaining")

            st.dataframe(
                fast_low,
                use_container_width=True,
                hide_index=True
            )

            st.write("### Official Ending Balance Rows")

            official_ending = df[df["Movement Type"] == "Ending Balance"]

            st.dataframe(
                official_ending,
                use_container_width=True,
                hide_index=True
            )

            st.write("### Official Total Rows")

            official_totals_check = df[df["Movement Type"] == "Total"]

            st.dataframe(
                official_totals_check,
                use_container_width=True,
                hide_index=True
            )

        # =========================
        # TAB 6: RECENT OUTBOUND AUDIT
        # =========================

        with tab6:
            st.subheader("Recent Outbound Audit")

            st.write("This tab shows exactly which rows are included in the 30D / 14D / 7D outbound calculations.")

            window_choice = st.radio(
                "Select outbound window",
                ["30D", "14D", "7D"],
                horizontal=True
            )

            audit_rows = recent_windows[window_choice]["rows"].copy()

            st.write(
                f"Window used: {recent_windows[window_choice]['start'].date()} "
                f"to {recent_windows[window_choice]['end'].date()}"
            )

            audit_summary = (
                audit_rows.groupby("SKU", as_index=False)[["Qty Out", "Qty Out CTN"]]
                .sum()
                .rename(columns={
                    "Qty Out": f"Outbound {window_choice}",
                    "Qty Out CTN": f"Outbound {window_choice} CTN"
                })
                .sort_values(f"Outbound {window_choice}", ascending=False)
            )

            st.write("### SKU-level summary for selected window")

            st.dataframe(
                audit_summary,
                use_container_width=True,
                hide_index=True
            )

            st.write("### Source transaction rows included")

            audit_cols = [
                "Source Row",
                "SKU",
                "Description",
                "Activity Date",
                "Activity Text",
                "Trans #",
                "Ref #",
                "Qty Out",
                "Qty Out CTN",
                "Balance",
                "Ctn Balance",
                "Is Not Shipped"
            ]

            st.dataframe(
                audit_rows[audit_cols].sort_values(["Activity Date", "SKU", "Source Row"]),
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                f"Download {window_choice} Audit Rows CSV",
                audit_rows[audit_cols].to_csv(index=False).encode("utf-8"),
                file_name=f"outbound_{window_choice}_audit_rows.csv",
                mime="text/csv"
            )

        # =========================
        # DOWNLOAD FULL DATA
        # =========================

        st.divider()

        d1, d2 = st.columns(2)

        with d1:
            st.download_button(
                "Download Full SKU Summary CSV",
                summary.to_csv(index=False).encode("utf-8"),
                file_name="inventory_summary.csv",
                mime="text/csv"
            )

        with d2:
            st.download_button(
                "Download Cleaned Transaction CSV",
                df.to_csv(index=False).encode("utf-8"),
                file_name="cleaned_transactions.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error("Something went wrong while processing the file.")
        st.write("Please check whether the uploaded file is the correct Item Activity Report format.")
        st.exception(e)
