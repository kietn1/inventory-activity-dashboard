import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
from datetime import timedelta


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Inventory Shortage Dashboard",
    page_icon="📦",
    layout="wide"
)


# ============================================================
# CSS STYLE
# ============================================================

st.markdown(
    """
    <style>
        .main-title {
            font-size: 34px;
            font-weight: 800;
            margin-bottom: 0px;
        }

        .sub-title {
            font-size: 15px;
            color: #666;
            margin-bottom: 20px;
        }

        .section-title {
            font-size: 22px;
            font-weight: 700;
            margin-top: 25px;
            margin-bottom: 10px;
        }

        .small-note {
            font-size: 13px;
            color: #777;
        }

        div[data-testid="stMetric"] {
            background-color: #ffffff;
            border: 1px solid #e6e6e6;
            padding: 14px;
            border-radius: 12px;
            box-shadow: 0px 1px 4px rgba(0,0,0,0.06);
        }

        div[data-testid="stMetricLabel"] {
            font-size: 13px;
            color: #555;
        }

        div[data-testid="stMetricValue"] {
            font-size: 26px;
            font-weight: 700;
        }

        .risk-critical {
            background-color: #fff1f0;
            border-left: 6px solid #d93025;
            padding: 12px;
            border-radius: 8px;
        }

        .risk-warning {
            background-color: #fff8e1;
            border-left: 6px solid #fbbc04;
            padding: 12px;
            border-radius: 8px;
        }

        .risk-watch {
            background-color: #e8f0fe;
            border-left: 6px solid #1a73e8;
            padding: 12px;
            border-radius: 8px;
        }

        .risk-safe {
            background-color: #e6f4ea;
            border-left: 6px solid #188038;
            padding: 12px;
            border-radius: 8px;
        }

        .block-container {
            padding-top: 2rem;
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HEADER
# ============================================================

st.markdown('<div class="main-title">📦 Inventory Shortage Dashboard</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Focus: shortage risk, recent outbound trend, ending balance, and forecast stockout date.</div>',
    unsafe_allow_html=True
)


# ============================================================
# HELPERS
# ============================================================

def normalize_cell(value):
    if pd.isna(value):
        return ""

    text = str(value)

    text = (
        text
        .replace("\xa0", " ")
        .replace("\t", " ")
        .replace(",", "")
        .strip()
    )

    text = re.sub(r"\s+", " ", text)
    return text


def clean_text(value):
    return normalize_cell(value)


def extract_number(value):
    """
    Extract unit quantity.

    Examples:
    5139 / 198        -> 5139
       15.0000 / 90  -> 15
    blank             -> 0
    """
    text = normalize_cell(value)

    if text == "":
        return 0.0

    if "/" in text:
        text = text.split("/")[0].strip()

    match = re.search(r"[-+]?\d*\.?\d+", text)

    if match:
        return float(match.group())

    return 0.0


def has_number(value):
    text = normalize_cell(value)
    return bool(re.search(r"[-+]?\d*\.?\d+", text))


def safe_row_text(row):
    values = []

    for value in row.tolist():
        values.append(normalize_cell(value).lower())

    return " ".join(values)


def parse_activity_date(activity_text):
    """
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


# ============================================================
# EXCEL STRUCTURE HELPERS
# ============================================================

def find_header_row(raw_df):
    for i in range(min(100, len(raw_df))):
        row_text = safe_row_text(raw_df.iloc[i])

        if "sku" in row_text and "activity date" in row_text:
            return i

    return None


def find_col(headers, contains_all=None, contains_any=None, fallback=None, exclude_any=None):
    contains_all = contains_all or []
    contains_any = contains_any or []
    exclude_any = exclude_any or []

    for idx, header in enumerate(headers):
        h = normalize_cell(header).lower()

        if not h:
            continue

        if any(ex in h for ex in exclude_any):
            continue

        if contains_all and not all(term in h for term in contains_all):
            continue

        if contains_any and not any(term in h for term in contains_any):
            continue

        return idx

    return fallback


def build_column_map(raw_df, header_row):
    headers = raw_df.iloc[header_row].tolist()
    headers_lower = [normalize_cell(x).lower() for x in headers]

    return {
        "sku": find_col(headers_lower, contains_any=["sku"], fallback=0),
        "description": find_col(headers_lower, contains_any=["description", "item description"], fallback=2),
        "activity_date": find_col(headers_lower, contains_all=["activity", "date"], fallback=7),
        "trans": find_col(headers_lower, contains_any=["trans"], fallback=9),
        "ref": find_col(headers_lower, contains_any=["ref"], fallback=10),
        "qty_in": find_col(
            headers_lower,
            contains_all=["qty", "in"],
            fallback=12,
            exclude_any=["out"]
        ),
        "qty_out": find_col(
            headers_lower,
            contains_all=["qty", "out"],
            fallback=14,
            exclude_any=["inbound"]
        ),
        "balance": find_col(
            headers_lower,
            contains_any=["balance"],
            fallback=19,
            exclude_any=["ctn"]
        )
    }


def get_cell(row, idx):
    if idx is None:
        return None

    if idx < 0 or idx >= len(row):
        return None

    return row.iloc[idx]


def read_numeric_with_fallback(row, main_idx, scan_start=None, scan_end=None):
    """
    Read main cell first.
    If blank, scan nearby cells.

    This handles cases where Qty cell appears shifted because of blank spaces.
    """
    main_value = get_cell(row, main_idx)
    main_text = normalize_cell(main_value)

    if main_text != "":
        return extract_number(main_value), main_text

    if scan_start is None:
        scan_start = max((main_idx or 0) - 1, 0)

    if scan_end is None:
        scan_end = min((main_idx or 0) + 3, len(row) - 1)

    scan_start = max(scan_start, 0)
    scan_end = min(scan_end, len(row) - 1)

    for col in range(scan_start, scan_end + 1):
        value = get_cell(row, col)
        text = normalize_cell(value)

        if text == "":
            continue

        if has_number(text):
            return extract_number(text), text

    return 0.0, ""


# ============================================================
# REPORT RANGE
# ============================================================

def extract_report_range(raw_df):
    report_start = None
    report_end = None

    for i in range(min(30, len(raw_df))):
        row = raw_df.iloc[i]
        row_text = safe_row_text(row)

        if "item activity from" in row_text:
            dates_found = []

            for value in row.tolist():
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


# ============================================================
# FILTERS / LOGIC
# ============================================================

def filter_recent_window(tx, end_date, days):
    """
    True calendar window.

    Last 30 days = report end date - 29 days through report end date.
    Not Shipped rows are included if Qty Out > 0.
    """
    start_date = end_date - timedelta(days=days - 1)

    filtered = tx[
        (tx["Activity Date"] >= start_date)
        & (tx["Activity Date"] <= end_date)
        & (tx["Qty Out"] > 0)
    ].copy()

    return filtered, start_date, end_date


def set_risk_filter(values):
    st.session_state["risk_filter_click"] = values


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


def format_days(value):
    if pd.isna(value):
        return ""

    return round(float(value), 1)


# ============================================================
# PARSER
# ============================================================

def parse_item_activity_report(uploaded_file):
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    report_start, report_end = extract_report_range(raw)
    header_row = find_header_row(raw)

    if header_row is None:
        raise ValueError("Cannot find Item Activity Report header row.")

    col = build_column_map(raw, header_row)

    records = []

    current_sku = None
    current_description = None

    for idx in range(header_row + 1, len(raw)):
        row = raw.iloc[idx]

        sku_cell = get_cell(row, col["sku"])
        description_cell = get_cell(row, col["description"])

        # New SKU block
        if pd.notna(sku_cell) and normalize_cell(sku_cell) != "":
            sku_text = normalize_cell(sku_cell)

            ignored_words = [
                "sku",
                "nan",
                "warehouse",
                "customer",
                "item activity",
                "activity date",
                "page",
                "total"
            ]

            if sku_text.lower() not in ignored_words:
                current_sku = sku_text
                current_description = clean_text(description_cell)

            continue

        activity_date_raw = get_cell(row, col["activity_date"])
        trans_no = get_cell(row, col["trans"])
        ref_no = get_cell(row, col["ref"])

        activity_text = clean_text(activity_date_raw)
        trans_text = clean_text(trans_no)
        ref_text = clean_text(ref_no)

        activity_text_lower = activity_text.lower()
        ref_text_lower = ref_text.lower()

        is_beginning_balance = activity_text_lower == "beginning balance"
        is_ending_balance = activity_text_lower == "ending balance"

        row_cells_lower = [normalize_cell(x).lower() for x in row.tolist()]
        is_total_row = (
            ref_text_lower == "total"
            or any(x == "total" for x in row_cells_lower)
        )

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

        qty_in_scan_start = col["qty_in"]
        qty_in_scan_end = max(col["qty_in"], col["qty_out"] - 1) if col["qty_out"] else col["qty_in"] + 2

        qty_out_scan_start = col["qty_out"]
        qty_out_scan_end = max(col["qty_out"], col["balance"] - 1) if col["balance"] else col["qty_out"] + 3

        qty_in_num, raw_qty_in = read_numeric_with_fallback(
            row,
            col["qty_in"],
            qty_in_scan_start,
            qty_in_scan_end
        )

        qty_out_num, raw_qty_out = read_numeric_with_fallback(
            row,
            col["qty_out"],
            qty_out_scan_start,
            qty_out_scan_end
        )

        balance_num, raw_balance = read_numeric_with_fallback(
            row,
            col["balance"],
            col["balance"],
            min(col["balance"] + 2, len(row) - 1)
        )

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
            "Activity Date": activity_date,
            "Activity Text": activity_text,
            "Trans #": trans_text,
            "Ref #": ref_text,
            "Raw Qty In": raw_qty_in,
            "Raw Qty Out": raw_qty_out,
            "Raw Balance": raw_balance,
            "Qty In": qty_in_num,
            "Qty Out": qty_out_num,
            "Balance": balance_num,
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


# ============================================================
# SUMMARY
# ============================================================

def build_summary(df, report_start=None, report_end=None):
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

    recent_end_date = report_max_date

    # Official ending balance
    ending_balance_rows = df[df["Movement Type"] == "Ending Balance"].copy()

    if not ending_balance_rows.empty:
        ending_balance_rows["Source Order"] = range(len(ending_balance_rows))

        latest_balance = (
            ending_balance_rows
            .sort_values(["SKU", "Source Order"])
            .groupby("SKU", as_index=False)
            .tail(1)[["SKU", "Description", "Balance"]]
            .rename(columns={"Balance": "Ending Balance"})
        )
    else:
        sorted_df = tx.copy()
        sorted_df["Source Order"] = range(len(sorted_df))

        latest_balance = (
            sorted_df
            .sort_values(["SKU", "Activity Date", "Source Order"])
            .groupby("SKU", as_index=False)
            .tail(1)[["SKU", "Description", "Balance"]]
            .rename(columns={"Balance": "Ending Balance"})
        )

    # Official totals
    total_rows = df[df["Movement Type"] == "Total"].copy()

    if not total_rows.empty:
        total_rows["Source Order"] = range(len(total_rows))

        official_totals = (
            total_rows
            .sort_values(["SKU", "Source Order"])
            .groupby("SKU", as_index=False)
            .tail(1)[["SKU", "Qty In", "Qty Out"]]
            .rename(columns={
                "Qty In": "Total Inbound",
                "Qty Out": "Total Outbound"
            })
        )

        total_inbound = official_totals[["SKU", "Total Inbound"]]
        total_outbound = official_totals[["SKU", "Total Outbound"]]

    else:
        total_inbound = (
            tx.groupby("SKU", as_index=False)["Qty In"]
            .sum()
            .rename(columns={"Qty In": "Total Inbound"})
        )

        total_outbound = (
            tx.groupby("SKU", as_index=False)["Qty Out"]
            .sum()
            .rename(columns={"Qty Out": "Total Outbound"})
        )

    last_activity = (
        tx.groupby("SKU", as_index=False)["Activity Date"]
        .max()
        .rename(columns={"Activity Date": "Last Activity Date"})
    )

    tx_30, outbound_30_start, outbound_30_end = filter_recent_window(tx, recent_end_date, 30)
    tx_14, outbound_14_start, outbound_14_end = filter_recent_window(tx, recent_end_date, 14)
    tx_7, outbound_7_start, outbound_7_end = filter_recent_window(tx, recent_end_date, 7)

    outbound_30 = (
        tx_30.groupby("SKU", as_index=False)["Qty Out"]
        .sum()
        .rename(columns={"Qty Out": "Outbound Last 30 Days"})
    )

    outbound_14 = (
        tx_14.groupby("SKU", as_index=False)["Qty Out"]
        .sum()
        .rename(columns={"Qty Out": "Outbound Last 14 Days"})
    )

    outbound_7 = (
        tx_7.groupby("SKU", as_index=False)["Qty Out"]
        .sum()
        .rename(columns={"Qty Out": "Outbound Last 7 Days"})
    )

    summary = latest_balance.copy()

    for small_df in [
        total_inbound,
        total_outbound,
        last_activity,
        outbound_30,
        outbound_14,
        outbound_7
    ]:
        summary = summary.merge(small_df, on="SKU", how="left")

    fill_cols = [
        "Total Inbound",
        "Total Outbound",
        "Outbound Last 30 Days",
        "Outbound Last 14 Days",
        "Outbound Last 7 Days"
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

    def recommended_action(row):
        risk = row["Risk Level"]

        if risk == "Critical":
            return "Immediate check / inbound follow-up"
        if risk == "Warning":
            return "Prepare replenishment"
        if risk == "Watch":
            return "Monitor closely"
        if risk == "No Recent Movement":
            return "No recent demand"
        return "OK"

    summary["Risk Level"] = summary.apply(risk_level, axis=1)
    summary["Recommended Action"] = summary.apply(recommended_action, axis=1)

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

    recent_windows = {
        "30D": {"start": outbound_30_start, "end": outbound_30_end, "rows": tx_30},
        "14D": {"start": outbound_14_start, "end": outbound_14_end, "rows": tx_14},
        "7D": {"start": outbound_7_start, "end": outbound_7_end, "rows": tx_7}
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


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("Upload")
uploaded_file = st.sidebar.file_uploader(
    "Upload Item Activity Report",
    type=["xlsx"]
)

st.sidebar.divider()
st.sidebar.header("Logic")
st.sidebar.write(
    """
    **Shortage forecast uses:**

    Ending Balance  
    Outbound Last 30 Days  
    Average Daily Usage  
    Days Remaining  

    Not Shipped rows are included when Qty Out > 0.
    """
)


# ============================================================
# MAIN APP
# ============================================================

if uploaded_file is None:
    st.info("Upload an Item Activity Report Excel file from the sidebar to start.")

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

        if "risk_filter_click" not in st.session_state:
            st.session_state["risk_filter_click"] = ["Critical", "Warning"]

        # ========================================================
        # TOP INFO
        # ========================================================

        st.markdown('<div class="section-title">Report Overview</div>', unsafe_allow_html=True)

        st.caption(
            f"Report Range: {report_min_date.date()} to {report_max_date.date()} | "
            f"Recent 30D Window: {recent_windows['30D']['start'].date()} to {recent_windows['30D']['end'].date()}"
        )

        total_skus = summary["SKU"].nunique()
        total_ending_balance = summary["Ending Balance"].sum()
        total_outbound_30 = summary["Outbound Last 30 Days"].sum()
        total_outbound_14 = summary["Outbound Last 14 Days"].sum()
        total_outbound_7 = summary["Outbound Last 7 Days"].sum()

        critical_count = (summary["Risk Level"] == "Critical").sum()
        warning_count = (summary["Risk Level"] == "Warning").sum()
        watch_count = (summary["Risk Level"] == "Watch").sum()

        k1, k2, k3, k4, k5, k6 = st.columns(6)

        with k1:
            st.metric("Total SKUs", f"{total_skus:,}")

        with k2:
            st.metric("Critical", f"{critical_count:,}")

        with k3:
            st.metric("Warning", f"{warning_count:,}")

        with k4:
            st.metric("Watch", f"{watch_count:,}")

        with k5:
            st.metric("Ending Balance", f"{total_ending_balance:,.0f}")

        with k6:
            st.metric("Outbound 30D", f"{total_outbound_30:,.0f}")

        # ========================================================
        # ACTION BUTTONS
        # ========================================================

        b1, b2, b3, b4 = st.columns(4)

        with b1:
            if st.button("View Critical", use_container_width=True):
                set_risk_filter(["Critical"])

        with b2:
            if st.button("View Warning", use_container_width=True):
                set_risk_filter(["Warning"])

        with b3:
            if st.button("View Critical + Warning", use_container_width=True):
                set_risk_filter(["Critical", "Warning"])

        with b4:
            if st.button("View All Risks", use_container_width=True):
                set_risk_filter(["Critical", "Warning", "Watch", "Safe", "No Recent Movement"])

        # ========================================================
        # RECENT OUTBOUND CARDS
        # ========================================================

        st.markdown('<div class="section-title">Recent Outbound</div>', unsafe_allow_html=True)

        r1, r2, r3 = st.columns(3)

        with r1:
            st.metric(
                "Last 30 Days",
                f"{total_outbound_30:,.0f}",
                help=f"{recent_windows['30D']['start'].date()} to {recent_windows['30D']['end'].date()}"
            )

        with r2:
            st.metric(
                "Last 14 Days",
                f"{total_outbound_14:,.0f}",
                help=f"{recent_windows['14D']['start'].date()} to {recent_windows['14D']['end'].date()}"
            )

        with r3:
            st.metric(
                "Last 7 Days",
                f"{total_outbound_7:,.0f}",
                help=f"{recent_windows['7D']['start'].date()} to {recent_windows['7D']['end'].date()}"
            )

        # ========================================================
        # MAIN SHORTAGE TABLE
        # ========================================================

        st.markdown('<div class="section-title">Shortage Priority List</div>', unsafe_allow_html=True)

        risk_filter = st.multiselect(
            "Risk filter",
            ["Critical", "Warning", "Watch", "Safe", "No Recent Movement"],
            default=st.session_state["risk_filter_click"]
        )

        shortage_view = summary[
            summary["Risk Level"].isin(risk_filter)
        ].copy()

        shortage_view = apply_risk_sort(shortage_view)

        shortage_view["Days Remaining"] = shortage_view["Days Remaining"].apply(format_days)

        priority_cols = [
            "SKU",
            "Description",
            "Risk Level",
            "Recommended Action",
            "Ending Balance",
            "Outbound Last 30 Days",
            "Outbound Last 14 Days",
            "Outbound Last 7 Days",
            "Avg Daily Usage 30D",
            "Days Remaining",
            "Forecast Stockout Date",
            "Last Activity Date"
        ]

        st.dataframe(
            shortage_view[priority_cols],
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "Download Shortage Priority CSV",
            shortage_view[priority_cols].to_csv(index=False).encode("utf-8"),
            file_name="shortage_priority.csv",
            mime="text/csv"
        )

        # ========================================================
        # TABS
        # ========================================================

        tab1, tab2, tab3 = st.tabs([
            "SKU Detail",
            "Trend",
            "Audit"
        ])

        # ========================================================
        # SKU DETAIL
        # ========================================================

        with tab1:
            st.subheader("SKU Detail")

            sku_list = sorted(summary["SKU"].dropna().unique())

            selected_sku = st.selectbox(
                "Select SKU",
                sku_list
            )

            sku_summary = summary[summary["SKU"] == selected_sku].copy()
            sku_summary["Days Remaining"] = sku_summary["Days Remaining"].apply(format_days)

            sku_summary_cols = [
                "SKU",
                "Description",
                "Risk Level",
                "Recommended Action",
                "Ending Balance",
                "Total Inbound",
                "Total Outbound",
                "Outbound Last 30 Days",
                "Outbound Last 14 Days",
                "Outbound Last 7 Days",
                "Avg Daily Usage 30D",
                "Days Remaining",
                "Forecast Stockout Date",
                "Last Activity Date"
            ]

            st.dataframe(
                sku_summary[sku_summary_cols],
                use_container_width=True,
                hide_index=True
            )

            sku_tx = tx[tx["SKU"] == selected_sku].sort_values("Activity Date")

            st.write("Transaction History")

            tx_cols = [
                "Source Row",
                "Activity Date",
                "Activity Text",
                "Trans #",
                "Ref #",
                "Raw Qty Out",
                "Qty In",
                "Qty Out",
                "Balance",
                "Is Not Shipped"
            ]

            st.dataframe(
                sku_tx[tx_cols],
                use_container_width=True,
                hide_index=True
            )

            if not sku_tx.empty:
                fig_sku = px.line(
                    sku_tx,
                    x="Activity Date",
                    y="Balance",
                    markers=True,
                    hover_data=["Trans #", "Ref #", "Qty In", "Qty Out"],
                    title=f"Balance Trend - {selected_sku}"
                )

                st.plotly_chart(fig_sku, use_container_width=True)

        # ========================================================
        # TREND
        # ========================================================

        with tab2:
            st.subheader("Daily Movement Trend")

            daily = (
                tx.groupby("Activity Date", as_index=False)[["Qty In", "Qty Out"]]
                .sum()
                .sort_values("Activity Date")
            )

            daily["Net Movement"] = daily["Qty In"] - daily["Qty Out"]

            fig_out = px.line(
                daily,
                x="Activity Date",
                y="Qty Out",
                markers=True,
                title="Daily Outbound"
            )

            st.plotly_chart(fig_out, use_container_width=True)

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

        # ========================================================
        # AUDIT
        # ========================================================

        with tab3:
            st.subheader("Audit")

            audit_choice = st.radio(
                "Audit view",
                [
                    "Recent Outbound 30D",
                    "Recent Outbound 14D",
                    "Recent Outbound 7D",
                    "Official Total Rows",
                    "Official Ending Balance Rows",
                    "Not Shipped Rows",
                    "Cancelled Transactions"
                ],
                horizontal=True
            )

            if audit_choice == "Recent Outbound 30D":
                audit_rows = recent_windows["30D"]["rows"]
                st.caption(
                    f"Window: {recent_windows['30D']['start'].date()} to {recent_windows['30D']['end'].date()}"
                )

            elif audit_choice == "Recent Outbound 14D":
                audit_rows = recent_windows["14D"]["rows"]
                st.caption(
                    f"Window: {recent_windows['14D']['start'].date()} to {recent_windows['14D']['end'].date()}"
                )

            elif audit_choice == "Recent Outbound 7D":
                audit_rows = recent_windows["7D"]["rows"]
                st.caption(
                    f"Window: {recent_windows['7D']['start'].date()} to {recent_windows['7D']['end'].date()}"
                )

            elif audit_choice == "Official Total Rows":
                audit_rows = df[df["Movement Type"] == "Total"]

            elif audit_choice == "Official Ending Balance Rows":
                audit_rows = df[df["Movement Type"] == "Ending Balance"]

            elif audit_choice == "Not Shipped Rows":
                audit_rows = df[df["Is Not Shipped"] == True]

            else:
                audit_rows = df[df["Is Cancelled"] == True]

            audit_cols = [
                "Source Row",
                "SKU",
                "Description",
                "Activity Date",
                "Activity Text",
                "Trans #",
                "Ref #",
                "Raw Qty In",
                "Raw Qty Out",
                "Raw Balance",
                "Qty In",
                "Qty Out",
                "Balance",
                "Movement Type",
                "Is Not Shipped"
            ]

            st.dataframe(
                audit_rows[audit_cols],
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                "Download Audit CSV",
                audit_rows[audit_cols].to_csv(index=False).encode("utf-8"),
                file_name="audit_rows.csv",
                mime="text/csv"
            )

        # ========================================================
        # FOOTER DOWNLOADS
        # ========================================================

        st.divider()

        d1, d2 = st.columns(2)

        with d1:
            st.download_button(
                "Download Full Summary CSV",
                summary.to_csv(index=False).encode("utf-8"),
                file_name="inventory_shortage_summary.csv",
                mime="text/csv"
            )

        with d2:
            st.download_button(
                "Download Cleaned Transactions CSV",
                df.to_csv(index=False).encode("utf-8"),
                file_name="cleaned_transactions.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error("Something went wrong while processing the file.")
        st.write("Please confirm the uploaded file is an Item Activity Report.")
        st.exception(e)
