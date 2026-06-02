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
    Convert values like:
    35.0000 / 35
    1.0000 / 1
    blank
    into clean number.
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


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def safe_row_text(row):
    """
    Convert every cell in a row to safe lowercase text.
    This prevents error:
    TypeError: sequence item expected str instance, float found
    """
    row_values = []

    for value in row.tolist():
        if pd.isna(value):
            row_values.append("")
        else:
            row_values.append(str(value).strip().lower())

    return " ".join(row_values)


def find_header_row(raw_df):
    """
    Find the header row automatically.
    We look for row containing both SKU and Activity Date.
    """
    for i in range(min(80, len(raw_df))):
        row_text = safe_row_text(raw_df.iloc[i])

        if "sku" in row_text and "activity date" in row_text:
            return i

    return None


def parse_item_activity_report(uploaded_file):
    """
    Main parser for Item Activity Report.

    Expected structure:
    - SKU appears once per item block
    - Transaction rows below SKU have blank SKU
    - Activity Date / Trans# / Ref# / Qty In / Qty Out / Balance
    """

    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

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

        if current_sku and pd.notna(activity_date_raw):
            activity_text = clean_text(activity_date_raw)

            if activity_text == "":
                continue

            is_beginning_balance = activity_text.lower() == "beginning balance"

            if is_beginning_balance:
                activity_date = pd.NaT
            else:
                activity_date = pd.to_datetime(activity_text, errors="coerce")

            qty_in_num = extract_number(qty_in)
            qty_out_num = extract_number(qty_out)
            balance_num = extract_number(balance)
            ctn_balance_num = extract_number(ctn_balance)

            if is_beginning_balance:
                movement_type = "Beginning Balance"
            elif qty_in_num > 0 and qty_out_num == 0:
                movement_type = "Inbound"
            elif qty_out_num > 0 and qty_in_num == 0:
                movement_type = "Outbound"
            elif qty_in_num > 0 and qty_out_num > 0:
                movement_type = "Mixed"
            else:
                movement_type = "No Movement"

            ref_text = clean_text(ref_no)
            trans_text = clean_text(trans_no)

            is_cancelled = (
                "cancel" in ref_text.lower()
                or "cancel" in trans_text.lower()
            )

            records.append({
                "SKU": current_sku,
                "Description": current_description,
                "Packed": current_packed,
                "Activity Date": activity_date,
                "Activity Text": activity_text,
                "Trans #": trans_text,
                "Ref #": ref_text,
                "Qty In": qty_in_num,
                "Qty Out": qty_out_num,
                "Balance": balance_num,
                "Ctn Balance": ctn_balance_num,
                "Movement Type": movement_type,
                "Is Cancelled": is_cancelled
            })

    df = pd.DataFrame(records)

    if df.empty:
        raise ValueError("No activity records found.")

    return df


def build_summary(df):
    """
    Build SKU-level summary.
    """

    tx = df[df["Activity Date"].notna()].copy()

    if tx.empty:
        raise ValueError("No valid transaction dates found.")

    min_date = tx["Activity Date"].min()
    max_date = tx["Activity Date"].max()

    sorted_df = df.copy()
    sorted_df["Sort Date"] = sorted_df["Activity Date"].fillna(pd.Timestamp.min)

    latest_balance = (
        sorted_df.sort_values(["SKU", "Sort Date"])
        .groupby("SKU", as_index=False)
        .tail(1)[["SKU", "Description", "Packed", "Balance", "Ctn Balance"]]
    )

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

    outbound_30 = (
        tx[tx["Activity Date"] >= max_date - timedelta(days=30)]
        .groupby("SKU", as_index=False)["Qty Out"]
        .sum()
        .rename(columns={"Qty Out": "Outbound Last 30 Days"})
    )

    outbound_14 = (
        tx[tx["Activity Date"] >= max_date - timedelta(days=14)]
        .groupby("SKU", as_index=False)["Qty Out"]
        .sum()
        .rename(columns={"Qty Out": "Outbound Last 14 Days"})
    )

    outbound_7 = (
        tx[tx["Activity Date"] >= max_date - timedelta(days=7)]
        .groupby("SKU", as_index=False)["Qty Out"]
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
        summary["Balance"] / summary["Avg Daily Usage 30D"],
        np.nan
    )

    def risk_level(row):
        balance = row["Balance"]
        usage = row["Avg Daily Usage 30D"]
        days = row["Days Remaining"]

        if balance <= 0 and usage > 0:
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
        return (max_date + timedelta(days=float(row["Days Remaining"]))).date()

    summary["Forecast Stockout Date"] = summary.apply(stockout_date, axis=1)

    summary["Report Start Date"] = min_date.date()
    summary["Report End Date"] = max_date.date()

    return summary, tx, min_date, max_date


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
        df = parse_item_activity_report(uploaded_file)
        summary, tx, min_date, max_date = build_summary(df)

        st.success(
            f"File processed successfully. Report range: {min_date.date()} to {max_date.date()}"
        )

        # =========================
        # SESSION STATE FOR INTERACTIVE KPI FILTER
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
        total_balance = summary["Balance"].sum()
        total_inbound = summary["Total Inbound"].sum()
        total_outbound = summary["Total Outbound"].sum()
        critical_count = (summary["Risk Level"] == "Critical").sum()
        warning_count = (summary["Risk Level"] == "Warning").sum()
        watch_count = (summary["Risk Level"] == "Watch").sum()
        safe_count = (summary["Risk Level"] == "Safe").sum()
        no_movement_count = (summary["Risk Level"] == "No Recent Movement").sum()

        k1, k2, k3, k4, k5, k6 = st.columns(6)

        with k1:
            st.metric("Total SKUs", f"{total_skus:,}")
            if st.button("View All SKUs", use_container_width=True):
                st.session_state["risk_filter_click"] = [
                    "Critical",
                    "Warning",
                    "Watch",
                    "Safe",
                    "No Recent Movement"
                ]

        with k2:
            st.metric("Current Balance", f"{total_balance:,.0f}")
            if st.button("View Safe SKUs", use_container_width=True):
                st.session_state["risk_filter_click"] = ["Safe"]

        with k3:
            st.metric("Total Inbound", f"{total_inbound:,.0f}")
            if st.button("View No Movement", use_container_width=True):
                st.session_state["risk_filter_click"] = ["No Recent Movement"]

        with k4:
            st.metric("Total Outbound", f"{total_outbound:,.0f}")
            if st.button("View Watch SKUs", use_container_width=True):
                st.session_state["risk_filter_click"] = ["Watch"]

        with k5:
            st.metric("Critical SKUs", f"{critical_count:,}")
            if st.button("View Critical SKUs", use_container_width=True):
                st.session_state["risk_filter_click"] = ["Critical"]

        with k6:
            st.metric("Warning SKUs", f"{warning_count:,}")
            if st.button("View Warning SKUs", use_container_width=True):
                st.session_state["risk_filter_click"] = ["Warning"]

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
            "Balance",
            "Total Outbound",
            "Outbound Last 30 Days",
            "Outbound Last 14 Days",
            "Outbound Last 7 Days",
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

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "Overview",
            "Shortage Forecast",
            "Movement Trend",
            "SKU Detail",
            "Exceptions"
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
                        "Balance",
                        "Risk Level"
                    ],
                    title="Top 15 Outbound SKUs"
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
                "Balance",
                "Total Inbound",
                "Total Outbound",
                "Outbound Last 30 Days",
                "Outbound Last 14 Days",
                "Outbound Last 7 Days",
                "Avg Daily Usage 30D",
                "Days Remaining",
                "Forecast Stockout Date",
                "Risk Level",
                "Last Activity Date"
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
                        "Qty Out"
                    ],
                    title=f"Balance Trend - {selected_sku}"
                )

                st.plotly_chart(fig_sku_balance, use_container_width=True)

        # =========================
        # TAB 5: EXCEPTIONS
        # =========================

        with tab5:
            st.subheader("Exception Report")

            st.write("### Cancelled Transactions")

            cancelled = df[df["Is Cancelled"] == True]

            st.dataframe(
                cancelled,
                use_container_width=True,
                hide_index=True
            )

            st.write("### Zero Balance with Recent Outbound")

            zero_recent = summary[
                (summary["Balance"] <= 0)
                & (summary["Outbound Last 30 Days"] > 0)
            ]

            st.dataframe(
                zero_recent,
                use_container_width=True,
                hide_index=True
            )

            st.write("### Fast Moving but Low Balance")

            fast_low = summary[
                (summary["Outbound Last 30 Days"] > 0)
                & (summary["Days Remaining"] <= 30)
            ].sort_values("Days Remaining")

            st.dataframe(
                fast_low,
                use_container_width=True,
                hide_index=True
            )

            st.write("### No Recent Movement")

            no_movement = summary[
                summary["Risk Level"] == "No Recent Movement"
            ]

            st.dataframe(
                no_movement,
                use_container_width=True,
                hide_index=True
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
