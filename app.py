import re
from datetime import date, datetime, timedelta
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from pandas.tseries.holiday import USFederalHolidayCalendar


# ============================================================
# Streamlit page setup
# ============================================================
st.set_page_config(
    page_title="Inventory Shortage Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .main .block-container {padding-top: 0.9rem; padding-bottom: 1.4rem; max-width: 1500px;}
        header, footer {visibility: hidden;}
        .kpi-card {
            border: 1px solid rgba(0,0,0,0.08);
            border-radius: 18px;
            padding: 16px 16px 13px 16px;
            background: rgba(255,255,255,0.86);
            box-shadow: 0 6px 24px rgba(16, 24, 40, 0.06);
            min-height: 104px;
        }
        .kpi-label {font-size: 0.82rem; color:#6B7280; font-weight: 650; margin-bottom: 7px;}
        .kpi-value {font-size: 1.75rem; color:#111827; font-weight: 800; line-height: 1.08; letter-spacing:-0.03em;}
        .kpi-help {font-size: 0.76rem; color:#9CA3AF; margin-top: 8px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
        .section-title {font-size:1.12rem; font-weight:800; color:#111827; margin-top: 0.3rem;}
        .section-subtitle {font-size:0.84rem; color:#6B7280; margin-bottom: 0.7rem;}
        .small-note {font-size:0.81rem; color:#6B7280;}
        div[data-testid="stDataFrame"] {border-radius: 14px; overflow: hidden;}
        div[data-testid="stSidebar"] {background:#F8FAFC;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# Utility functions
# ============================================================
FIXED_COLS = {
    "sku": 0,
    "description": 2,
    "activity_date": 7,
    "trans_no": 9,
    "ref_no": 10,
    "qty_in": 12,
    "qty_out": 14,
    "balance": 19,
}


def clean_text(value) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    return str(value).replace("\u00a0", " ").replace("\u200b", "").strip()


def first_qty_number(value) -> float:
    """Parse Qty like ' 5,139 / 198' and return first number only: 5139."""
    text = clean_text(value)
    if not text:
        return 0.0
    text = text.replace(",", "")
    before_slash = text.split("/")[0]
    match = re.search(r"-?\d+(?:\.\d+)?", before_slash)
    return float(match.group()) if match else 0.0


def parse_excel_or_text_date(value):
    """Parse dates including '6/1/2026 (Not Shipped)' and Excel serial dates."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return pd.NaT

    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.to_datetime(value).normalize()

    if isinstance(value, date):
        return pd.to_datetime(value).normalize()

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel serial dates are usually > 30000 for modern dates.
        if value > 30000:
            return pd.to_datetime(value, unit="D", origin="1899-12-30").normalize()
        return pd.NaT

    text = clean_text(value)
    if not text:
        return pd.NaT

    # Remove status note but keep the date part.
    text = re.sub(r"\s*\([^)]*\)", "", text).strip()
    parsed = pd.to_datetime(text, errors="coerce")
    return parsed.normalize() if not pd.isna(parsed) else pd.NaT


def find_header_row(raw: pd.DataFrame) -> int:
    for idx in range(len(raw)):
        row_values = [clean_text(x).lower() for x in raw.iloc[idx].tolist()]
        if "sku" in row_values and "activity date" in row_values and "ref #" in row_values:
            return idx
    raise ValueError("Không tìm thấy header row có SKU / Activity Date / Ref #.")


def extract_report_range(raw: pd.DataFrame):
    start_dt, end_dt = pd.NaT, pd.NaT
    for r in range(min(15, len(raw))):
        row = raw.iloc[r].tolist()
        row_text = " | ".join(clean_text(x) for x in row).lower()
        if "item activity from" in row_text:
            # Start = first date-like value after label. End = first date-like value after 'to'.
            to_index = None
            for i, value in enumerate(row):
                if clean_text(value).lower() == "to":
                    to_index = i
                    break

            dates_before_to, dates_after_to = [], []
            for i, value in enumerate(row):
                parsed = parse_excel_or_text_date(value)
                if not pd.isna(parsed):
                    if to_index is not None and i > to_index:
                        dates_after_to.append(parsed)
                    else:
                        dates_before_to.append(parsed)

            if dates_before_to:
                start_dt = dates_before_to[0]
            if dates_after_to:
                end_dt = dates_after_to[0]
            elif len(dates_before_to) >= 2:
                end_dt = dates_before_to[1]
            break

    return start_dt, end_dt


@st.cache_data(show_spinner=False)
def load_excel_to_raw(file_bytes: bytes) -> pd.DataFrame:
    # header=None is important because the file has report title rows above the actual header.
    return pd.read_excel(BytesIO(file_bytes), sheet_name=0, header=None, dtype=object)


@st.cache_data(show_spinner=False)
def process_excel_file(file_bytes: bytes) -> dict:
    raw_df = load_excel_to_raw(file_bytes)
    return build_inventory_model(raw_df)



def us_holidays_around(end_date: pd.Timestamp, lookback_days: int = 730) -> set:
    """Return US federal holidays around the report end date for working-day windows."""
    end_date = pd.to_datetime(end_date).normalize()
    cal = USFederalHolidayCalendar()
    holidays = cal.holidays(
        start=end_date - pd.Timedelta(days=lookback_days),
        end=end_date + pd.Timedelta(days=lookback_days),
    )
    return {pd.to_datetime(x).normalize() for x in holidays}


def is_valid_working_day(day, holidays: set) -> bool:
    """Valid working day = Monday-Friday and not a US federal holiday."""
    day = pd.to_datetime(day).normalize()
    return day.weekday() < 5 and day not in holidays


def last_valid_working_dates(end_date, count: int, holidays: set) -> list:
    """Get exactly the last N valid working dates, including report end date if valid."""
    dates = []
    current = pd.to_datetime(end_date).normalize()
    while len(dates) < count:
        if is_valid_working_day(current, holidays):
            dates.append(current)
        current -= pd.Timedelta(days=1)
    return sorted(dates)


def add_valid_working_days(start_date, days, holidays: set):
    """Add floor(days) valid working days, excluding Saturday/Sunday/holidays."""
    if pd.isna(days) or not np.isfinite(days):
        return pd.NaT
    remaining = int(np.floor(days))
    current = pd.to_datetime(start_date).normalize()
    while remaining > 0:
        current += pd.Timedelta(days=1)
        if is_valid_working_day(current, holidays):
            remaining -= 1
    return current


def build_inventory_model(raw: pd.DataFrame) -> dict:
    header_idx = find_header_row(raw)
    report_start, report_end = extract_report_range(raw)

    rows = raw.iloc[header_idx + 1 :].copy()

    current_sku = ""
    current_desc = ""
    sku_records = {}
    transactions = []
    official_total_rows = []
    official_ending_rows = []
    not_shipped_rows = []
    cancelled_rows = []

    for excel_row_num, row in rows.iterrows():
        sku_cell = clean_text(row.iloc[FIXED_COLS["sku"]] if len(row) > FIXED_COLS["sku"] else "")
        desc_cell = clean_text(row.iloc[FIXED_COLS["description"]] if len(row) > FIXED_COLS["description"] else "")
        activity_raw = row.iloc[FIXED_COLS["activity_date"]] if len(row) > FIXED_COLS["activity_date"] else None
        activity_text = clean_text(activity_raw)
        ref_text = clean_text(row.iloc[FIXED_COLS["ref_no"]] if len(row) > FIXED_COLS["ref_no"] else "")
        trans_no = clean_text(row.iloc[FIXED_COLS["trans_no"]] if len(row) > FIXED_COLS["trans_no"] else "")
        qty_in = first_qty_number(row.iloc[FIXED_COLS["qty_in"]] if len(row) > FIXED_COLS["qty_in"] else None)
        qty_out = first_qty_number(row.iloc[FIXED_COLS["qty_out"]] if len(row) > FIXED_COLS["qty_out"] else None)
        balance = first_qty_number(row.iloc[FIXED_COLS["balance"]] if len(row) > FIXED_COLS["balance"] else None)

        # SKU section row: SKU appears in column A, usually no real transaction date/ref.
        if sku_cell and sku_cell.lower() != "sku" and ref_text.lower() != "total":
            current_sku = sku_cell
            if desc_cell:
                current_desc = desc_cell
            sku_records.setdefault(
                current_sku,
                {
                    "SKU": current_sku,
                    "Description": current_desc,
                    "Official Total Inbound": 0.0,
                    "Official Total Outbound": 0.0,
                    "Ending Balance": 0.0,
                    "Official Ending Row": None,
                    "Official Total Row": None,
                    "Last Activity Date": pd.NaT,
                },
            )
            if desc_cell:
                sku_records[current_sku]["Description"] = desc_cell

        if not current_sku:
            continue

        sku_records.setdefault(
            current_sku,
            {
                "SKU": current_sku,
                "Description": current_desc,
                "Official Total Inbound": 0.0,
                "Official Total Outbound": 0.0,
                "Ending Balance": 0.0,
                "Official Ending Row": None,
                "Official Total Row": None,
                "Last Activity Date": pd.NaT,
            },
        )

        if activity_text.lower() == "ending balance":
            sku_records[current_sku]["Ending Balance"] = balance
            sku_records[current_sku]["Official Ending Row"] = excel_row_num + 1
            official_ending_rows.append(
                {
                    "Excel Row": excel_row_num + 1,
                    "SKU": current_sku,
                    "Description": sku_records[current_sku]["Description"],
                    "Activity Date": activity_text,
                    "Balance": balance,
                }
            )
            continue

        # Official total row must be Ref # = Total, not Activity Date.
        if ref_text.lower() == "total":
            sku_records[current_sku]["Official Total Inbound"] = qty_in
            sku_records[current_sku]["Official Total Outbound"] = qty_out
            sku_records[current_sku]["Official Total Row"] = excel_row_num + 1
            official_total_rows.append(
                {
                    "Excel Row": excel_row_num + 1,
                    "SKU": current_sku,
                    "Description": sku_records[current_sku]["Description"],
                    "Ref #": ref_text,
                    "Official Total Inbound": qty_in,
                    "Official Total Outbound": qty_out,
                    "Balance": balance,
                }
            )
            continue

        activity_dt = parse_excel_or_text_date(activity_raw)
        is_cancelled = "cancel" in ref_text.lower()
        is_not_shipped = "not shipped" in activity_text.lower()

        if is_cancelled:
            cancelled_rows.append(
                {
                    "Excel Row": excel_row_num + 1,
                    "SKU": current_sku,
                    "Description": sku_records[current_sku]["Description"],
                    "Activity Date Raw": activity_text,
                    "Activity Date": activity_dt,
                    "Ref #": ref_text,
                    "Qty Out": qty_out,
                }
            )

        if is_not_shipped:
            not_shipped_rows.append(
                {
                    "Excel Row": excel_row_num + 1,
                    "SKU": current_sku,
                    "Description": sku_records[current_sku]["Description"],
                    "Activity Date Raw": activity_text,
                    "Parsed Date": activity_dt,
                    "Ref #": ref_text,
                    "Qty Out": qty_out,
                }
            )

        # Dated transaction rows for recent outbound. Count Not Shipped if Qty Out > 0.
        if not pd.isna(activity_dt) and qty_out > 0:
            transactions.append(
                {
                    "Excel Row": excel_row_num + 1,
                    "SKU": current_sku,
                    "Description": sku_records[current_sku]["Description"],
                    "Activity Date Raw": activity_text,
                    "Activity Date": activity_dt,
                    "Trans. #": trans_no,
                    "Ref #": ref_text,
                    "Qty Out": qty_out,
                    "Qty In": qty_in,
                    "Balance After Transaction": balance,
                    "Is Not Shipped": is_not_shipped,
                    "Is Cancelled": is_cancelled,
                }
            )

            existing_last = sku_records[current_sku]["Last Activity Date"]
            if pd.isna(existing_last) or activity_dt > existing_last:
                sku_records[current_sku]["Last Activity Date"] = activity_dt

    sku_df = pd.DataFrame(sku_records.values())
    tx_df = pd.DataFrame(transactions)
    official_total_df = pd.DataFrame(official_total_rows)
    official_ending_df = pd.DataFrame(official_ending_rows)
    not_shipped_df = pd.DataFrame(not_shipped_rows)
    cancelled_df = pd.DataFrame(cancelled_rows)

    if pd.isna(report_end):
        # Fallback: use max activity date if header range cannot be parsed.
        report_end = tx_df["Activity Date"].max() if not tx_df.empty else pd.Timestamp.today().normalize()
    if pd.isna(report_start):
        report_start = tx_df["Activity Date"].min() if not tx_df.empty else report_end

    report_end = pd.to_datetime(report_end).normalize()
    report_start = pd.to_datetime(report_start).normalize()

    holidays = us_holidays_around(report_end)
    window_dates = {
        "Outbound Last 30 Days": last_valid_working_dates(report_end, 30, holidays),
        "Outbound Last 14 Days": last_valid_working_dates(report_end, 14, holidays),
        "Outbound Last 7 Days": last_valid_working_dates(report_end, 7, holidays),
    }
    windows = {label: (dates[0], dates[-1]) for label, dates in window_dates.items()}

    if sku_df.empty:
        raise ValueError("Không tìm thấy SKU nào trong file.")

    for label, dates in window_dates.items():
        if tx_df.empty:
            sku_df[label] = 0.0
        else:
            mask = tx_df["Activity Date"].isin(dates)
            agg = tx_df.loc[mask].groupby("SKU", as_index=True)["Qty Out"].sum()
            sku_df[label] = sku_df["SKU"].map(agg).fillna(0.0)

    sku_df["Avg Daily Usage 30D"] = sku_df["Outbound Last 30 Days"] / len(window_dates["Outbound Last 30 Days"])
    sku_df["Days Remaining"] = np.where(
        sku_df["Avg Daily Usage 30D"] > 0,
        sku_df["Ending Balance"] / sku_df["Avg Daily Usage 30D"],
        np.inf,
    )

    demand_exists = (sku_df["Outbound Last 30 Days"] > 0) | (sku_df["Outbound Last 14 Days"] > 0) | (sku_df["Outbound Last 7 Days"] > 0)
    conditions = [
        (sku_df["Ending Balance"] <= 0) & demand_exists,
        (sku_df["Outbound Last 7 Days"] > 0) & (sku_df["Days Remaining"] <= 7),
        (sku_df["Outbound Last 14 Days"] > 0) & (sku_df["Days Remaining"] <= 14),
        (sku_df["Outbound Last 30 Days"] > 0) & (sku_df["Days Remaining"] <= 30),
    ]
    sku_df["Risk Level"] = np.select(conditions, ["Critical", "Critical", "Warning", "Watch"], default="Healthy")
    sku_df["Recommended Action"] = sku_df["Risk Level"].map({
        "Critical": "Prepare inbound / allocate stock immediately",
        "Warning": "Review inbound ETA and reserve inventory",
        "Watch": "Monitor weekly usage and upcoming orders",
        "Healthy": "No immediate action",
    })

    def stockout_date(days_remaining):
        if pd.isna(days_remaining) or not np.isfinite(days_remaining):
            return pd.NaT
        if days_remaining > 365:
            return pd.NaT
        return add_valid_working_days(report_end, days_remaining, holidays)

    sku_df["Forecast Stockout Date"] = sku_df["Days Remaining"].map(stockout_date)

    risk_order = {"Critical": 0, "Warning": 1, "Watch": 2, "Healthy": 3}
    sku_df["Risk Sort"] = sku_df["Risk Level"].map(risk_order).fillna(9)
    sku_df = sku_df.sort_values(
        by=["Risk Sort", "Days Remaining", "Outbound Last 14 Days", "Outbound Last 30 Days"],
        ascending=[True, True, False, False],
    ).reset_index(drop=True)

    # Daily trend for outbound.
    if tx_df.empty:
        trend_df = pd.DataFrame(columns=["Activity Date", "Qty Out"])
    else:
        trend_df = tx_df.groupby("Activity Date", as_index=False)["Qty Out"].sum().sort_values("Activity Date")

    return {
        "sku_df": sku_df,
        "tx_df": tx_df,
        "trend_df": trend_df,
        "official_total_df": official_total_df,
        "official_ending_df": official_ending_df,
        "not_shipped_df": not_shipped_df,
        "cancelled_df": cancelled_df,
        "report_start": report_start,
        "report_end": report_end,
        "windows": windows,
        "window_dates": window_dates,
        "holidays": holidays,
        "header_idx": header_idx,
    }


def fmt_num(value, decimals=0):
    if value is None or pd.isna(value):
        return "-"
    if value == np.inf:
        return "∞"
    return f"{value:,.{decimals}f}"


def fmt_date(value):
    if value is None or pd.isna(value):
        return "-"
    return pd.to_datetime(value).strftime("%m/%d/%Y")


def risk_badge_text(level: str) -> str:
    return {
        "Critical": "🔴 Critical",
        "Warning": "🟠 Warning",
        "Watch": "🟡 Watch",
        "Healthy": "🟢 Healthy",
    }.get(level, level)


def metric_card(label, value, help_text=""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-help">{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def round_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    numeric_cols = out.select_dtypes(include=["number"]).columns
    for col in numeric_cols:
        if str(col) == "Avg Daily Usage 30D":
            out[col] = out[col].round(2)
        else:
            out[col] = out[col].round(0)
    return out


def format_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    date_keywords = ("date", "activity date", "forecast stockout", "last activity")
    for col in out.columns:
        col_lower = str(col).lower()
        if any(keyword in col_lower for keyword in date_keywords):
            parsed = pd.to_datetime(out[col], errors="coerce")
            if parsed.notna().any():
                out[col] = parsed.dt.strftime("%m/%d/%Y").replace("NaT", "")
    return out


def display_table(df: pd.DataFrame) -> pd.DataFrame:
    return format_date_columns(round_numeric_columns(df))


def prepare_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Risk Level"] = out["Risk Level"].map(risk_badge_text)
    integer_metric_cols = [
        "Ending Balance",
        "Official Total Inbound",
        "Official Total Outbound",
        "Outbound Last 30 Days",
        "Outbound Last 14 Days",
        "Outbound Last 7 Days",
        "Days Remaining",
        "Official Total Row",
        "Official Ending Row",
    ]
    for c in integer_metric_cols:
        if c in out.columns:
            out[c] = out[c].replace(np.inf, np.nan).round(0).astype("Int64")
    if "Avg Daily Usage 30D" in out.columns:
        out["Avg Daily Usage 30D"] = out["Avg Daily Usage 30D"].round(2)
    for c in ["Forecast Stockout Date", "Last Activity Date"]:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce").dt.strftime("%m/%d/%Y").replace("NaT", "")
    return out


@st.cache_data(show_spinner=False)
def to_excel_bytes(model: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        display_table(prepare_display(model["sku_df"]).drop(columns=["Risk Sort"], errors="ignore")).to_excel(writer, sheet_name="Shortage Priority", index=False)
        display_table(model["tx_df"]).to_excel(writer, sheet_name="Dated Transactions", index=False)
        display_table(model["official_total_df"]).to_excel(writer, sheet_name="Official Total Rows", index=False)
        display_table(model["official_ending_df"]).to_excel(writer, sheet_name="Ending Balance Rows", index=False)
        display_table(model["not_shipped_df"]).to_excel(writer, sheet_name="Not Shipped Rows", index=False)
        display_table(model["cancelled_df"]).to_excel(writer, sheet_name="Cancelled Rows", index=False)

        for worksheet in writer.sheets.values():
            for row in worksheet.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                        if worksheet.cell(row=1, column=cell.column).value == 'Avg Daily Usage 30D':
                            cell.number_format = '#,##0.00'
                        else:
                            cell.number_format = '#,##0'
    return output.getvalue()


def show_limited_dataframe(df: pd.DataFrame, height: int = 420, limit: int = 500):
    """Render only the first N rows for faster browser performance."""
    total_rows = len(df)
    if total_rows > limit:
        st.caption(f"Showing first {limit:,} rows out of {total_rows:,} rows for faster loading. Download the export for full data.")
    else:
        st.caption(f"Showing {total_rows:,} rows.")
    st.dataframe(display_table(df.head(limit)), use_container_width=True, hide_index=True, height=height)


# ============================================================
# Sidebar controls
# ============================================================
st.sidebar.title("Inventory Dashboard")
st.sidebar.caption("Fast shortage view")

uploaded = st.sidebar.file_uploader(
    "Drop Excel file here",
    type=["xlsx", "xls"],
    help="File format tương tự Item Activity Report.",
)

st.sidebar.divider()
st.sidebar.subheader("Risk Filter")
with st.sidebar.form("filter_form"):
    show_risks = st.multiselect(
        "Risk Level",
        options=["Critical", "Warning", "Watch", "Healthy"],
        default=["Critical", "Warning", "Watch"],
    )

    min_usage = st.number_input("Minimum Outbound Last 30 Days", min_value=0, value=0, step=1)
    search_text = st.text_input("Search SKU / Description", placeholder="Example: SBED, BACKUP SWITCH...")
    st.form_submit_button("Apply Filters")

st.sidebar.divider()
st.sidebar.markdown("""
#### Risk Level Notes

- **Critical:** 0–7 days remaining
- **Warning:** 8–14 days remaining
- **Watch:** 15–30 days remaining
- **Healthy:** More than 30 days remaining
- **No Recent Demand:** Outbound 30D = 0

Recent windows include the report date and count backward by valid working days only, excluding Saturdays, Sundays, and US federal holidays. Not Shipped and Cancelled rows are still counted when Qty Out > 0.
""")

# ============================================================
# Main app
# ============================================================
st.title("Inventory Shortage")
st.caption("Fast, clean shortage dashboard for Item Activity Report.")

if uploaded is None:
    st.info("Upload an Item Activity Report Excel file from the left sidebar to generate the dashboard.")
    st.stop()

try:
    file_bytes = uploaded.getvalue()
    model = process_excel_file(file_bytes)
except Exception as exc:
    st.error("File could not be processed. Please check if this is the correct Item Activity Report format.")
    st.exception(exc)
    st.stop()

sku_df = model["sku_df"].copy()
filtered = sku_df.copy()
if show_risks:
    filtered = filtered[filtered["Risk Level"].isin(show_risks)]
filtered = filtered[filtered["Outbound Last 30 Days"] >= min_usage]
if search_text.strip():
    q = search_text.strip().lower()
    filtered = filtered[
        filtered["SKU"].astype(str).str.lower().str.contains(q, na=False)
        | filtered["Description"].astype(str).str.lower().str.contains(q, na=False)
    ]

report_start = model["report_start"]
report_end = model["report_end"]
windows = model["windows"]

st.markdown(
    f"<div class='small-note'>Report Range: <b>{fmt_date(report_start)}</b> to <b>{fmt_date(report_end)}</b> | Recent windows use valid working days, including the report date.</div>",
    unsafe_allow_html=True,
)

# KPI cards
critical_count = int((sku_df["Risk Level"] == "Critical").sum())
warning_count = int((sku_df["Risk Level"] == "Warning").sum())
watch_count = int((sku_df["Risk Level"] == "Watch").sum())
healthy_count = int((sku_df["Risk Level"] == "Healthy").sum())

k1, k2, k3, k4 = st.columns(4)
with k1:
    metric_card("Total SKUs", fmt_num(len(sku_df)), f"Healthy: {healthy_count:,}")
with k2:
    metric_card("Critical SKUs", fmt_num(critical_count), "Need immediate inventory action")
with k3:
    metric_card("Warning SKUs", fmt_num(warning_count), "Need ETA / reserve review")
with k4:
    metric_card("Watch SKUs", fmt_num(watch_count), "Monitor usage trend")

k5, k6, k7, k8 = st.columns(4)
with k5:
    metric_card("Ending Balance", fmt_num(sku_df["Ending Balance"].sum()), "From official Ending Balance rows")
with k6:
    metric_card("Official Total Outbound", fmt_num(sku_df["Official Total Outbound"].sum()), "From official Ref # = Total rows")
with k7:
    metric_card("Recent Outbound 30D", fmt_num(sku_df["Outbound Last 30 Days"].sum()), f"{fmt_date(windows['Outbound Last 30 Days'][0])} - {fmt_date(windows['Outbound Last 30 Days'][1])}")
with k8:
    metric_card("Recent Outbound 14D / 7D", f"{fmt_num(sku_df['Outbound Last 14 Days'].sum())} / {fmt_num(sku_df['Outbound Last 7 Days'].sum())}", "Dated Qty Out rows only")

st.markdown("<div class='section-title'>Shortage Priority List</div>", unsafe_allow_html=True)
st.markdown("<div class='section-subtitle'>Sorted by risk level, lowest days remaining, and recent outbound demand.</div>", unsafe_allow_html=True)

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
    "Last Activity Date",
]

priority_display = prepare_display(filtered[priority_cols])
show_limited_dataframe(priority_display, height=440, limit=250)

with st.expander("Export", expanded=False):
    st.caption("Excel export is generated only when this section is opened, so the dashboard stays fast while filtering.")
    st.download_button(
        "⬇️ Download processed shortage report",
        data=to_excel_bytes(model),
        file_name=f"shortage_dashboard_export_{report_end.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# Tabs
sku_tab, trend_tab, audit_tab = st.tabs(["SKU Detail", "Trend", "Audit"])

with sku_tab:
    left, right = st.columns([1, 2])
    with left:
        sku_options = filtered["SKU"].tolist() if not filtered.empty else sku_df["SKU"].tolist()
        selected_sku = st.selectbox("Select SKU", options=sku_options)
    selected = sku_df[sku_df["SKU"] == selected_sku].iloc[0]

    d1, d2, d3, d4 = st.columns(4)
    with d1:
        metric_card("Risk Level", risk_badge_text(selected["Risk Level"]), selected["Recommended Action"])
    with d2:
        metric_card("Ending Balance", fmt_num(selected["Ending Balance"]), "Official Ending Balance")
    with d3:
        metric_card("Days Remaining", fmt_num(selected["Days Remaining"]), "Based on Avg Daily Usage 30D")
    with d4:
        metric_card("Forecast Stockout", fmt_date(selected["Forecast Stockout Date"]), "Forecast from report end date")

    st.subheader(f"{selected_sku} — {selected['Description']}")
    detail_cols = [
        "Official Total Inbound",
        "Official Total Outbound",
        "Outbound Last 30 Days",
        "Outbound Last 14 Days",
        "Outbound Last 7 Days",
        "Avg Daily Usage 30D",
        "Last Activity Date",
        "Official Total Row",
        "Official Ending Row",
    ]
    detail = selected[detail_cols].to_frame("Value")
    detail["Value"] = detail.apply(
        lambda r: fmt_date(r["Value"]) if "date" in str(r.name).lower()
        else (
            fmt_num(r["Value"], 2) if str(r.name) == "Avg Daily Usage 30D"
            else (fmt_num(r["Value"]) if isinstance(r["Value"], (int, float, np.integer, np.floating)) and np.isfinite(r["Value"]) else r["Value"])
        ),
        axis=1,
    )
    st.dataframe(detail, use_container_width=True)

    tx_sku = model["tx_df"]
    if not tx_sku.empty:
        tx_sku = tx_sku[tx_sku["SKU"] == selected_sku].sort_values("Activity Date", ascending=False)
        st.subheader("Dated outbound transactions")
        show_limited_dataframe(tx_sku, height=340, limit=250)

with trend_tab:
    st.subheader("Outbound Trend")
    trend_df = model["trend_df"]
    if trend_df.empty:
        st.info("No dated outbound transactions found.")
    else:
        trend_plot = trend_df.copy()
        trend_plot["Activity Date"] = pd.to_datetime(trend_plot["Activity Date"])
        st.line_chart(trend_plot, x="Activity Date", y="Qty Out", height=360, use_container_width=True)

        top_usage = sku_df.sort_values("Outbound Last 30 Days", ascending=False).head(20)[["SKU", "Outbound Last 30 Days"]]
        st.bar_chart(top_usage, x="SKU", y="Outbound Last 30 Days", height=360, use_container_width=True)

with audit_tab:
    st.subheader("Audit Checks")
    st.caption("Use this tab to verify recent outbound calculations and official source rows.")

    # Recent outbound summary cards
    st.markdown("**Recent Outbound Audit**")
    r1, r2, r3 = st.columns(3)
    recent_labels = ["Outbound Last 30 Days", "Outbound Last 14 Days", "Outbound Last 7 Days"]
    recent_card_labels = ["Recent Outbound 30D", "Recent Outbound 14D", "Recent Outbound 7D"]
    for col, label, card_label in zip([r1, r2, r3], recent_labels, recent_card_labels):
        start, end = windows[label]
        valid_dates = model["window_dates"][label]
        with col:
            metric_card(
                card_label,
                fmt_num(sku_df[label].sum()),
                f"{fmt_date(start)} - {fmt_date(end)} | {len(valid_dates)} working days",
            )

    st.markdown("**Source Row Audit**")
    s1, s2, s3, s4 = st.columns(4)
    with s1:
        metric_card("Official Total Rows", fmt_num(len(model["official_total_df"])), "Ref # = Total")
    with s2:
        metric_card("Ending Balance Rows", fmt_num(len(model["official_ending_df"])), "Activity Date = Ending Balance")
    with s3:
        metric_card("Not Shipped Rows", fmt_num(len(model["not_shipped_df"])), "Counted if Qty Out > 0")
    with s4:
        metric_card("Cancelled Transactions", fmt_num(len(model["cancelled_df"])), "Counted if Qty Out > 0")

    st.divider()

    audit_group_col, audit_table_col = st.columns([1, 2])
    with audit_group_col:
        audit_group = st.selectbox(
            "Audit Category",
            ["Recent Outbound", "Official Source Rows", "Status Rows"],
        )

    if audit_group == "Recent Outbound":
        audit_options = ["Recent Outbound 30D", "Recent Outbound 14D", "Recent Outbound 7D"]
    elif audit_group == "Official Source Rows":
        audit_options = ["Official Total Rows", "Official Ending Balance Rows"]
    else:
        audit_options = ["Not Shipped Rows", "Cancelled Transactions"]

    with audit_table_col:
        audit_choice = st.selectbox("Audit Table", audit_options)

    if audit_choice.startswith("Recent Outbound"):
        label_map = {
            "Recent Outbound 30D": "Outbound Last 30 Days",
            "Recent Outbound 14D": "Outbound Last 14 Days",
            "Recent Outbound 7D": "Outbound Last 7 Days",
        }
        label = label_map[audit_choice]
        start, end = windows[label]
        tx = model["tx_df"]
        if tx.empty:
            st.info("No dated outbound transactions found.")
        else:
            valid_dates = model["window_dates"][label]
            audit_tx = tx[tx["Activity Date"].isin(valid_dates)].copy()
            st.write(
                f"Window: **{fmt_date(start)} – {fmt_date(end)}** | "
                f"Valid working dates counted: **{len(valid_dates)}** | "
                f"Rows counted: **{len(audit_tx):,}** | "
                f"Qty Out total: **{fmt_num(audit_tx['Qty Out'].sum())}**"
            )
            show_limited_dataframe(audit_tx.sort_values(["SKU", "Activity Date"]), height=420)
    elif audit_choice == "Official Total Rows":
        show_limited_dataframe(model["official_total_df"], height=420)
    elif audit_choice == "Official Ending Balance Rows":
        show_limited_dataframe(model["official_ending_df"], height=420)
    elif audit_choice == "Not Shipped Rows":
        show_limited_dataframe(model["not_shipped_df"], height=420)
    elif audit_choice == "Cancelled Transactions":
        show_limited_dataframe(model["cancelled_df"], height=420)
