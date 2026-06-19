"""Delivery Fill Rate Dashboard — SAP/BW export analysis and reporting."""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

from src.fill_rate_engine import (
    build_kpis,
    build_plant_summary,
    build_product_group_summary,
    build_product_summary,
    build_shortage_report,
    build_top10,
    generate_excel_report,
    load_fill_rate,
)
from src.column_variants import REPORT_DELIVERY_SHORTAGE, apply_columns
from src.column_variants_ui import render_variant_panel

st.title("Delivery Fill Rate Dashboard")
st.caption(
    "Upload a SAP/BW Delivery Fill Rate Excel export to analyze shortages, "
    "fill rates by plant and product, and download a polished Excel report."
)

# ---------------------------------------------------------------------------
# Sidebar — upload
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Upload & Settings")
    uploaded_file = st.file_uploader(
        "SAP/BW Delivery Fill Rate Export (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="dfr_upload",
    )
    threshold = st.number_input(
        "High Priority Dollar Threshold ($)",
        min_value=0.0,
        value=1000.0,
        step=100.0,
        help="Lines with Short Amount >= this value are marked High Priority.",
    )
    st.markdown("---")
    st.caption(
        "**Note:** The app automatically detects the real header row, "
        "skips technical/blank rows, and fills down SAP merged-cell values."
    )

if uploaded_file is None:
    st.info("Upload a SAP/BW Delivery Fill Rate Excel file using the sidebar to get started.")
    st.stop()

# ---------------------------------------------------------------------------
# Load & process (cached per file + threshold)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _process(file_bytes: bytes, threshold: float):
    return load_fill_rate(io.BytesIO(file_bytes), threshold)


with st.spinner("Reading and cleaning data — please wait..."):
    try:
        df_clean, df_raw = _process(uploaded_file.getvalue(), threshold)
    except Exception as exc:
        st.error(f"Could not read file: {exc}")
        st.stop()

if df_clean.empty:
    st.warning("The file was read but no data rows were found.")
    st.stop()

# ---------------------------------------------------------------------------
# Filter panel (sidebar, below upload)
# ---------------------------------------------------------------------------
# Use a session-state counter to reset widget keys on demand.
if "dfr_filter_version" not in st.session_state:
    st.session_state["dfr_filter_version"] = 0

_v = st.session_state["dfr_filter_version"]

with st.sidebar:
    st.markdown("---")
    st.header("Filters")

    def _opts(col: str) -> list:
        if col not in df_clean.columns:
            return []
        return sorted(df_clean[col].dropna().unique().tolist())

    sel_plant = st.multiselect("Plant", _opts("plant"), key=f"dfr_plant_{_v}")
    sel_product = st.multiselect("Product", _opts("product"), key=f"dfr_product_{_v}")
    sel_pgroup = st.multiselect("Product Group", _opts("product_group_name"), key=f"dfr_pgroup_{_v}")
    sel_obd = st.multiselect("Outbound Delivery", _opts("outbound_delivery"), key=f"dfr_obd_{_v}")
    sel_so = st.multiselect("Sales Order", _opts("sales_order"), key=f"dfr_so_{_v}")

    # Date range filter
    has_date = "requested_delivery_date" in df_clean.columns
    if has_date:
        valid_dates = df_clean["requested_delivery_date"].dropna()
        min_date = valid_dates.min().date() if not valid_dates.empty else date.today()
        max_date = valid_dates.max().date() if not valid_dates.empty else date.today()
        date_range = st.date_input(
            "Requested Delivery Date",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            key=f"dfr_date_{_v}",
        )
        if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
            sel_date_start, sel_date_end = date_range
        else:
            sel_date_start = sel_date_end = None
    else:
        sel_date_start = sel_date_end = None

    sel_status = st.multiselect("Shortage Status", _opts("Shortage Status"), key=f"dfr_status_{_v}")
    sel_priority = st.multiselect("Priority", _opts("Priority"), key=f"dfr_priority_{_v}")
    sel_wh = st.multiselect("WH Fill Status", _opts("WH Fill Status"), key=f"dfr_wh_{_v}")
    sel_cust = st.multiselect("Customer Fill Status", _opts("Customer Fill Status"), key=f"dfr_cust_{_v}")

    st.markdown("---")
    sort_options = [
        "Highest Short Amount",
        "Highest Short Qty",
        "Lowest WH Fill Rate",
        "Lowest Customer Fill Rate",
        "Requested Delivery Date",
        "Product",
        "Plant",
    ]
    sort_by = st.selectbox("Sort By", sort_options, index=0, key=f"dfr_sort_{_v}")

    if st.button("Reset Filters", key="dfr_reset"):
        st.session_state["dfr_filter_version"] += 1
        st.rerun()

# ---------------------------------------------------------------------------
# Apply filters to df_clean → df_filtered
# ---------------------------------------------------------------------------
df_filtered = df_clean.copy()

if sel_plant:
    df_filtered = df_filtered[df_filtered["plant"].isin(sel_plant)]
if sel_product:
    df_filtered = df_filtered[df_filtered["product"].isin(sel_product)]
if sel_pgroup and "product_group_name" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["product_group_name"].isin(sel_pgroup)]
if sel_obd and "outbound_delivery" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["outbound_delivery"].isin(sel_obd)]
if sel_so and "sales_order" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["sales_order"].isin(sel_so)]
if has_date and sel_date_start and sel_date_end:
    mask = (
        df_filtered["requested_delivery_date"].dt.date >= sel_date_start
    ) & (
        df_filtered["requested_delivery_date"].dt.date <= sel_date_end
    )
    df_filtered = df_filtered[mask]
if sel_status and "Shortage Status" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["Shortage Status"].isin(sel_status)]
if sel_priority and "Priority" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["Priority"].isin(sel_priority)]
if sel_wh and "WH Fill Status" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["WH Fill Status"].isin(sel_wh)]
if sel_cust and "Customer Fill Status" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["Customer Fill Status"].isin(sel_cust)]

df_filtered = df_filtered.reset_index(drop=True)

# ---------------------------------------------------------------------------
# Apply sort to df_filtered
# ---------------------------------------------------------------------------
_short_amt_col = "total_short_amount" if "total_short_amount" in df_filtered.columns else "short_amount"

_sort_map = {
    "Highest Short Amount":      (_short_amt_col,           False),
    "Highest Short Qty":         ("short_quantity",          False),
    "Lowest WH Fill Rate":       ("wh_fill_rate",            True),
    "Lowest Customer Fill Rate": ("customer_fill_rate",      True),
    "Requested Delivery Date":   ("requested_delivery_date", True),
    "Product":                   ("product",                 True),
    "Plant":                     ("plant",                   True),
}

_sort_col, _sort_asc = _sort_map.get(sort_by, (_short_amt_col, False))
if _sort_col in df_filtered.columns:
    df_filtered = df_filtered.sort_values(_sort_col, ascending=_sort_asc, na_position="last").reset_index(drop=True)

# ---------------------------------------------------------------------------
# Rebuild all summaries from filtered data
# ---------------------------------------------------------------------------
kpis       = build_kpis(df_filtered)
shortage_df = build_shortage_report(df_filtered)
# Shortage report sorted by highest short amount by default
if _short_amt_col in shortage_df.columns:
    shortage_df = shortage_df.sort_values(_short_amt_col, ascending=False).reset_index(drop=True)

product_df = build_product_summary(df_filtered)
# Product summary (Column C = product) sorted by Short Amount (Column N) descending
if product_df is not None and not product_df.empty:
    amt_cols_prod = [c for c in product_df.columns if "amount" in c.lower()]
    if amt_cols_prod:
        product_df = product_df.sort_values(amt_cols_prod[0], ascending=False).reset_index(drop=True)

plant_df   = build_plant_summary(df_filtered)
product_group_df = build_product_group_summary(df_filtered)
# Product group summary sorted by highest Short Amount
if product_group_df is not None and not product_group_df.empty:
    amt_cols_pg = [c for c in product_group_df.columns if "amount" in c.lower()]
    if amt_cols_pg:
        product_group_df = product_group_df.sort_values(amt_cols_pg[0], ascending=False).reset_index(drop=True)
top10      = build_top10(df_filtered)

# Active filter indicator
_active_filters = any([
    sel_plant, sel_product, sel_pgroup, sel_obd, sel_so,
    sel_status, sel_priority, sel_wh, sel_cust,
    (has_date and sel_date_start and sel_date_end and
     (sel_date_start != min_date if has_date else False)),
])
if _active_filters:
    st.info(
        f"Filters active — showing **{len(df_filtered):,}** of **{len(df_clean):,}** rows. "
        "Use **Reset Filters** in the sidebar to clear."
    )
    st.success(
        f"**{kpis['num_shorted_lines']:,}** shorted lines · "
        f"**{kpis['num_products_impacted']:,}** products impacted"
    )
else:
    st.success(
        f"Loaded **{len(df_clean):,}** rows — "
        f"**{kpis['num_shorted_lines']:,}** shorted lines, "
        f"**{kpis['num_products_impacted']:,}** products impacted."
    )

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
(
    tab_exec,
    tab_short,
    tab_prod,
    tab_pgroup,
    tab_plant,
    tab_top10,
    tab_raw,
    tab_download,
) = st.tabs([
    "Executive Dashboard",
    "Shortage Report",
    "Product Summary",
    "Product Group Summary",
    "Plant Summary",
    "Top 10 Issues",
    "Raw Preview",
    "Download Report",
])

# ── Executive Dashboard ─────────────────────────────────────────────────────
with tab_exec:
    st.subheader("Key Performance Indicators")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Order Qty",       f"{kpis['total_order_qty']:,.0f}")
    c2.metric("Total Delivered Qty",   f"{kpis['total_delivered_qty']:,.0f}")
    c3.metric("Total Short Qty",       f"{kpis['total_short_qty']:,.0f}")

    c4, c5, c6 = st.columns(3)
    c4.metric("Total Short Amount",        f"${kpis['total_short_amount']:,.2f}")
    c5.metric("Overall WH Fill Rate",      f"{kpis['overall_wh_fill_rate']:.1f}%")
    c6.metric("Overall Customer Fill Rate",f"{kpis['overall_customer_fill_rate']:.1f}%")

    c7, c8, c9 = st.columns(3)
    c7.metric("Outbound Deliveries", f"{kpis['num_deliveries']:,}")
    c8.metric("Shorted Lines",       f"{kpis['num_shorted_lines']:,}")
    c9.metric("Products Impacted",   f"{kpis['num_products_impacted']:,}")

    st.markdown("---")
    st.subheader("Charts")

    col_l, col_r = st.columns(2)

    if "Shortage Status" in df_filtered.columns:
        status_counts = df_filtered["Shortage Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]
        fig_pie = px.pie(
            status_counts,
            names="Status",
            values="Count",
            title="Shorted vs Fully Delivered",
            color="Status",
            color_discrete_map={"Shorted": "#EF4444", "Fully Delivered": "#22C55E"},
            hole=0.4,
        )
        fig_pie.update_layout(margin=dict(t=40, b=0, l=0, r=0))
        col_l.plotly_chart(fig_pie, use_container_width=True)

    if plant_df is not None and not plant_df.empty:
        plant_col  = plant_df.columns[0]
        rate_cols  = [c for c in plant_df.columns if "fill rate" in c.lower()]
        if rate_cols:
            fig_plant_rate = px.bar(
                plant_df.sort_values(rate_cols[0]),
                x=rate_cols[0],
                y=plant_col,
                orientation="h",
                title="WH Fill Rate by Plant",
                labels={rate_cols[0]: "Avg WH Fill Rate (%)", plant_col: "Plant"},
                color=rate_cols[0],
                color_continuous_scale=["#EF4444", "#F59E0B", "#22C55E"],
                range_color=[0, 100],
            )
            fig_plant_rate.update_layout(
                margin=dict(t=40, b=0, l=0, r=0),
                coloraxis_showscale=False,
                yaxis=dict(autorange="reversed"),
            )
            col_r.plotly_chart(fig_plant_rate, use_container_width=True)

    col_l2, col_r2 = st.columns(2)

    if plant_df is not None and not plant_df.empty:
        plant_col = plant_df.columns[0]
        amt_cols  = [c for c in plant_df.columns if "amount" in c.lower()]
        if amt_cols:
            fig_plant_amt = px.bar(
                plant_df.sort_values(amt_cols[0], ascending=False),
                x=plant_col,
                y=amt_cols[0],
                title="Short Amount by Plant",
                labels={amt_cols[0]: "Short Amount ($)", plant_col: "Plant"},
                color=amt_cols[0],
                color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig_plant_amt.update_layout(
                margin=dict(t=40, b=0, l=0, r=0),
                coloraxis_showscale=False,
            )
            col_l2.plotly_chart(fig_plant_amt, use_container_width=True)

    if product_df is not None and not product_df.empty:
        prod_col  = product_df.columns[0]
        qty_cols  = [c for c in product_df.columns if "short qty" in c.lower()]
        if qty_cols:
            top_prods = product_df.sort_values(qty_cols[0], ascending=False).head(15)
            fig_prod_qty = px.bar(
                top_prods,
                x=qty_cols[0],
                y=prod_col,
                orientation="h",
                title="Top 15 Products by Short Qty",
                labels={qty_cols[0]: "Short Qty", prod_col: "Product"},
                color=qty_cols[0],
                color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig_prod_qty.update_layout(
                margin=dict(t=40, b=0, l=0, r=0),
                coloraxis_showscale=False,
                yaxis=dict(autorange="reversed"),
            )
            col_r2.plotly_chart(fig_prod_qty, use_container_width=True)

# ── Shortage Report ─────────────────────────────────────────────────────────
with tab_short:
    st.subheader(f"Shortage Report — {len(shortage_df):,} problem lines")
    st.caption("Sorted by highest Short Amount. Use sidebar filters to narrow results.")

    shortage_cols = render_variant_panel(
        REPORT_DELIVERY_SHORTAGE,
        list(shortage_df.columns),
        key_prefix="dfr_shortage_variant",
    )

    if shortage_df.empty:
        st.success("No shortages or fill rate issues found.")
    else:
        st.dataframe(
            apply_columns(shortage_df, shortage_cols),
            use_container_width=True,
            hide_index=True,
        )

# ── Product Summary ─────────────────────────────────────────────────────────
with tab_prod:
    st.subheader("Product Summary")
    st.caption("Grouped by Product (Column C) · sorted by highest Short Amount (Column N)")
    if product_df is None or product_df.empty:
        st.info("Product column not found or no data to summarize.")
    else:
        st.dataframe(product_df, use_container_width=True, hide_index=True)
        prod_col  = product_df.columns[0]
        amt_cols  = [c for c in product_df.columns if "amount" in c.lower()]
        if amt_cols:
            top15 = product_df.head(15)
            fig = px.bar(
                top15,
                x=prod_col,
                y=amt_cols[0],
                title="Top 15 Products by Short Amount",
                labels={amt_cols[0]: "Short Amount ($)", prod_col: "Product"},
                color=amt_cols[0],
                color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ── Product Group Summary ────────────────────────────────────────────────────
with tab_pgroup:
    st.subheader("Product Group Summary")
    st.caption("Grouped by Product Group · sorted by highest Short Amount.")
    if product_group_df is None or product_group_df.empty:
        st.info("Product Group column not found or no data to summarize.")
    else:
        st.dataframe(product_group_df, use_container_width=True, hide_index=True)
        pg_col   = product_group_df.columns[0]
        amt_cols = [c for c in product_group_df.columns if "amount" in c.lower()]
        if amt_cols:
            top15 = product_group_df.head(15)
            fig = px.bar(
                top15,
                x=pg_col,
                y=amt_cols[0],
                title="Top 15 Product Groups by Short Amount",
                labels={amt_cols[0]: "Short Amount ($)", pg_col: "Product Group"},
                color=amt_cols[0],
                color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ── Plant Summary ────────────────────────────────────────────────────────────
with tab_plant:
    st.subheader("Plant Summary")
    if plant_df is None or plant_df.empty:
        st.info("Plant column not found or no data to summarize.")
    else:
        st.dataframe(plant_df, use_container_width=True, hide_index=True)
        plant_col = plant_df.columns[0]
        amt_cols  = [c for c in plant_df.columns if "amount" in c.lower()]
        if amt_cols:
            fig = px.bar(
                plant_df.sort_values(amt_cols[0], ascending=False),
                x=plant_col,
                y=amt_cols[0],
                title="Short Amount by Plant",
                labels={amt_cols[0]: "Short Amount ($)", plant_col: "Plant"},
                color=amt_cols[0],
                color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ── Top 10 Issues ────────────────────────────────────────────────────────────
with tab_top10:
    st.subheader("Top 10 Issues")
    sections = [
        ("Top 10 Products by Short Amount",        top10.get("products_by_short_amount")),
        ("Top 10 Products by Short Qty",           top10.get("products_by_short_qty")),
        ("Top 10 Product Groups by Short Amount",  top10.get("product_groups_by_short_amount")),
        ("Top 10 Plants by Short Amount",          top10.get("plants_by_short_amount")),
        ("Top 10 Deliveries by Short Amount",      top10.get("deliveries_by_short_amount")),
    ]
    col_a, col_b = st.columns(2)
    for idx, (title, t_df) in enumerate(sections):
        col = col_a if idx % 2 == 0 else col_b
        col.markdown(f"**{title}**")
        if t_df is not None and not t_df.empty:
            col.dataframe(t_df, use_container_width=True, hide_index=True)
            val_col = t_df.columns[1]
            fig = px.bar(
                t_df,
                x=val_col,
                y=t_df.columns[0],
                orientation="h",
                labels={val_col: val_col, t_df.columns[0]: ""},
                color=val_col,
                color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(
                margin=dict(t=10, b=0, l=0, r=0),
                coloraxis_showscale=False,
                showlegend=False,
                yaxis=dict(autorange="reversed"),
                height=280,
            )
            col.plotly_chart(fig, use_container_width=True)
        else:
            col.info("No data available.")

# ── Raw Preview ──────────────────────────────────────────────────────────────
with tab_raw:
    st.subheader("Raw Export Preview")
    st.caption("First 500 rows of the file as read — before any cleaning.")
    st.dataframe(df_raw.head(500), use_container_width=True, hide_index=True)

    st.subheader("Cleaned Data Preview")
    st.caption(
        "Filtered + sorted data after header detection, type conversion, "
        "fill-down, and calculated columns."
    )
    st.dataframe(df_filtered.head(500), use_container_width=True, hide_index=True)

# ── Download Report ──────────────────────────────────────────────────────────
with tab_download:
    st.subheader("Download Excel Report")
    st.markdown(
        "Generates a fully formatted Excel workbook with 9 sheets: "
        "Instructions, Executive Summary, Clean Data, Shortage Report, "
        "Product Summary, Plant Summary, Product Group Summary, Top 10 Issues, "
        "and Raw Export Preview."
    )
    if _active_filters:
        st.info(
            f"The report will reflect your active filters ({len(df_filtered):,} of "
            f"{len(df_clean):,} rows). Clear filters in the sidebar to export all data."
        )

    if st.button("Generate Excel Report", type="primary"):
        with st.spinner("Building Excel workbook..."):
            excel_bytes = generate_excel_report(
                df_clean=df_filtered,
                df_raw=df_raw,
                kpis=kpis,
                shortage_df=apply_columns(shortage_df, shortage_cols),
                product_df=product_df if product_df is not None else pd.DataFrame(),
                plant_df=plant_df if plant_df is not None else pd.DataFrame(),
                top10=top10,
                threshold=threshold,
                product_group_df=product_group_df if product_group_df is not None else pd.DataFrame(),
            )
        st.download_button(
            label="Download Delivery_Fill_Rate_Report.xlsx",
            data=excel_bytes,
            file_name="Delivery_Fill_Rate_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
