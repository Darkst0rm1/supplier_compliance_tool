"""Delivery Fill Rate Dashboard — SAP/BW export analysis and reporting."""
from __future__ import annotations

import io

import pandas as pd
import plotly.express as px
import streamlit as st

from src.fill_rate_engine import (
    build_kpis,
    build_plant_summary,
    build_product_summary,
    build_shortage_report,
    build_top10,
    generate_excel_report,
    load_fill_rate,
)

st.set_page_config(page_title="Delivery Fill Rate Dashboard", layout="wide")
st.title("Delivery Fill Rate Dashboard")
st.caption(
    "Upload a SAP/BW Delivery Fill Rate Excel export to analyze shortages, "
    "fill rates by plant and product, and download a polished Excel report."
)

# ---------------------------------------------------------------------------
# Sidebar — inputs
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
# Load & process data (cached per file + threshold)
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
    st.warning("The file was read but no data rows were found. Check that the file contains actual data.")
    st.stop()

kpis = build_kpis(df_clean)
shortage_df = build_shortage_report(df_clean)
product_df = build_product_summary(df_clean)
plant_df = build_plant_summary(df_clean)
top10 = build_top10(df_clean)

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
    tab_plant,
    tab_top10,
    tab_raw,
    tab_download,
) = st.tabs([
    "Executive Dashboard",
    "Shortage Report",
    "Product Summary",
    "Plant Summary",
    "Top 10 Issues",
    "Raw Preview",
    "Download Report",
])

# ── Executive Dashboard ─────────────────────────────────────────────────────
with tab_exec:
    st.subheader("Key Performance Indicators")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Order Qty", f"{kpis['total_order_qty']:,.0f}")
    c2.metric("Total Delivered Qty", f"{kpis['total_delivered_qty']:,.0f}")
    c3.metric("Total Short Qty", f"{kpis['total_short_qty']:,.0f}")

    c4, c5, c6 = st.columns(3)
    c4.metric("Total Short Amount", f"${kpis['total_short_amount']:,.2f}")
    c5.metric("Overall WH Fill Rate", f"{kpis['overall_wh_fill_rate']:.1f}%")
    c6.metric("Overall Customer Fill Rate", f"{kpis['overall_customer_fill_rate']:.1f}%")

    c7, c8, c9 = st.columns(3)
    c7.metric("Outbound Deliveries", f"{kpis['num_deliveries']:,}")
    c8.metric("Shorted Lines", f"{kpis['num_shorted_lines']:,}")
    c9.metric("Products Impacted", f"{kpis['num_products_impacted']:,}")

    st.markdown("---")
    st.subheader("Charts")

    col_l, col_r = st.columns(2)

    # Shorted vs Fully Delivered
    if "Shortage Status" in df_clean.columns:
        status_counts = df_clean["Shortage Status"].value_counts().reset_index()
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

    # Fill Rate by Plant
    if plant_df is not None and not plant_df.empty:
        plant_col = plant_df.columns[0]
        rate_cols = [c for c in plant_df.columns if "fill rate" in c.lower()]
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

    # Short Amount by Plant
    if plant_df is not None and not plant_df.empty:
        plant_col = plant_df.columns[0]
        amt_cols = [c for c in plant_df.columns if "amount" in c.lower()]
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

    # Short Quantity by Product (top 15)
    if product_df is not None and not product_df.empty:
        prod_col = product_df.columns[0]
        qty_cols = [c for c in product_df.columns if "short qty" in c.lower()]
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
    if shortage_df.empty:
        st.success("No shortages or fill rate issues found.")
    else:
        if "Priority" in shortage_df.columns:
            priorities = ["All"] + sorted(shortage_df["Priority"].dropna().unique().tolist())
            sel_priority = st.selectbox("Filter by Priority", priorities, key="short_priority")
            view = shortage_df if sel_priority == "All" else shortage_df[shortage_df["Priority"] == sel_priority]
        else:
            view = shortage_df
        st.dataframe(view, use_container_width=True, hide_index=True)

# ── Product Summary ─────────────────────────────────────────────────────────
with tab_prod:
    st.subheader("Product Summary")
    if product_df is None or product_df.empty:
        st.info("Product column not found or no data to summarize.")
    else:
        st.dataframe(product_df, use_container_width=True, hide_index=True)
        prod_col = product_df.columns[0]
        amt_cols = [c for c in product_df.columns if "amount" in c.lower()]
        if amt_cols:
            top15 = product_df.sort_values(amt_cols[0], ascending=False).head(15)
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

# ── Plant Summary ────────────────────────────────────────────────────────────
with tab_plant:
    st.subheader("Plant Summary")
    if plant_df is None or plant_df.empty:
        st.info("Plant column not found or no data to summarize.")
    else:
        st.dataframe(plant_df, use_container_width=True, hide_index=True)
        plant_col = plant_df.columns[0]
        amt_cols = [c for c in plant_df.columns if "amount" in c.lower()]
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
        ("Top 10 Products by Short Amount", top10.get("products_by_short_amount")),
        ("Top 10 Products by Short Qty", top10.get("products_by_short_qty")),
        ("Top 10 Plants by Short Amount", top10.get("plants_by_short_amount")),
        ("Top 10 Deliveries by Short Amount", top10.get("deliveries_by_short_amount")),
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
    st.caption("Data after header detection, type conversion, fill-down, and calculated columns.")
    st.dataframe(df_clean.head(500), use_container_width=True, hide_index=True)

# ── Download Report ──────────────────────────────────────────────────────────
with tab_download:
    st.subheader("Download Excel Report")
    st.markdown(
        "Generates a fully formatted Excel workbook with 8 sheets: "
        "Instructions, Executive Summary, Clean Data, Shortage Report, "
        "Product Summary, Plant Summary, Top 10 Issues, and Raw Export Preview."
    )

    if st.button("Generate Excel Report", type="primary"):
        with st.spinner("Building Excel workbook..."):
            excel_bytes = generate_excel_report(
                df_clean=df_clean,
                df_raw=df_raw,
                kpis=kpis,
                shortage_df=shortage_df,
                product_df=product_df if product_df is not None else pd.DataFrame(),
                plant_df=plant_df if plant_df is not None else pd.DataFrame(),
                top10=top10,
                threshold=threshold,
            )
        st.download_button(
            label="Download Delivery_Fill_Rate_Report.xlsx",
            data=excel_bytes,
            file_name="Delivery_Fill_Rate_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
