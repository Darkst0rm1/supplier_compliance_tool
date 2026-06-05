"""Sales Order Fill Rate Dashboard — SAP/BW export analysis and reporting."""
from __future__ import annotations

import io

import pandas as pd
import plotly.express as px
import streamlit as st

from src.sales_order_engine import (
    build_account_summary,
    build_cdm_summary,
    build_kpis,
    build_plant_summary,
    build_product_summary,
    build_top10,
    build_unconfirmed_report,
    generate_excel_report,
    load_sales_order,
)

st.title("Sales Order Fill Rate Dashboard")
st.caption(
    "Upload a SAP/BW Sales Order Fill Rate export to analyze unconfirmed demand, "
    "fill rates by account and product, and download a polished Excel report."
)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Upload & Settings")
    uploaded_file = st.file_uploader(
        "SAP/BW Sales Order Fill Rate Export (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="so_upload",
    )
    threshold = st.number_input(
        "High Priority Dollar Threshold ($)",
        min_value=0.0,
        value=500.0,
        step=100.0,
        help="Lines with Unconfirmed Demand Amount >= this value are marked High Priority.",
    )
    st.markdown("---")
    st.caption(
        "**Note:** The app automatically detects the real header row and "
        "skips SAP technical rows above the data."
    )

if uploaded_file is None:
    st.info("Upload a SAP/BW Sales Order Fill Rate Excel file using the sidebar to get started.")
    st.stop()

# ---------------------------------------------------------------------------
# Load & process
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _process(file_bytes: bytes, threshold: float):
    return load_sales_order(io.BytesIO(file_bytes), threshold)


with st.spinner("Reading and cleaning data — please wait..."):
    try:
        df_clean, df_raw = _process(uploaded_file.getvalue(), threshold)
    except Exception as exc:
        st.error(f"Could not read file: {exc}")
        st.stop()

if df_clean.empty:
    st.warning("No data rows found. Check that the file contains actual data.")
    st.stop()

kpis           = build_kpis(df_clean)
unconfirmed_df = build_unconfirmed_report(df_clean)
account_df     = build_account_summary(df_clean)
product_df     = build_product_summary(df_clean)
plant_df       = build_plant_summary(df_clean)
cdm_df         = build_cdm_summary(df_clean)
top10          = build_top10(df_clean)

st.success(
    f"Loaded **{len(df_clean):,}** rows — "
    f"**{kpis['num_unconfirmed_lines']:,}** lines with unconfirmed demand — "
    f"**{kpis['num_key_accounts']:,}** key accounts."
)

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
(
    tab_exec, tab_unc, tab_acct, tab_prod,
    tab_plant, tab_cdm, tab_top10, tab_raw, tab_dl,
) = st.tabs([
    "Executive Dashboard",
    "Unconfirmed Demand Report",
    "Key Account Summary",
    "Product Summary",
    "Plant Summary",
    "CDM Summary",
    "Top 10 Issues",
    "Raw Preview",
    "Download Report",
])

# ── Executive Dashboard ──────────────────────────────────────────────────────
with tab_exec:
    st.subheader("Key Performance Indicators")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Order Qty",       f"{kpis['total_order_qty']:,.0f}")
    c2.metric("Total Confirmed Qty",   f"{kpis['total_confirmed_qty']:,.0f}")
    c3.metric("Total Unconfirmed Qty", f"{kpis['total_unconfirmed_qty']:,.0f}")

    c4, c5, c6 = st.columns(3)
    c4.metric("Total Net Value",              f"${kpis['total_net_value']:,.2f}")
    c5.metric("Total Unconfirmed Demand ($)", f"${kpis['total_unconfirmed_amount']:,.2f}")
    c6.metric("Overall Fill Rate",            f"{kpis['overall_fill_rate']:.1f}%")

    c7, c8, c9, c10 = st.columns(4)
    c7.metric("Sales Orders",               f"{kpis['num_sales_orders']:,}")
    c8.metric("Key Accounts",               f"{kpis['num_key_accounts']:,}")
    c9.metric("Products Impacted",          f"{kpis['num_products_impacted']:,}")
    c10.metric("Unconfirmed Demand Lines",  f"{kpis['num_unconfirmed_lines']:,}")

    st.markdown("---")
    st.subheader("Charts")

    col_l, col_r = st.columns(2)

    # Demand Status breakdown
    if "Demand Status" in df_clean.columns:
        status_counts = df_clean["Demand Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]
        fig = px.pie(
            status_counts, names="Status", values="Count",
            title="Demand Status Breakdown",
            color="Status",
            color_discrete_map={
                "Unconfirmed Demand": "#EF4444",
                "Fully Confirmed":    "#22C55E",
            },
            hole=0.4,
        )
        fig.update_layout(margin=dict(t=40, b=0, l=0, r=0))
        col_l.plotly_chart(fig, use_container_width=True)

    # Priority breakdown
    if "Priority" in df_clean.columns:
        pri_counts = df_clean["Priority"].value_counts().reset_index()
        pri_counts.columns = ["Priority", "Count"]
        fig = px.bar(
            pri_counts, x="Priority", y="Count",
            title="Priority Breakdown",
            color="Priority",
            color_discrete_map={
                "High Priority":   "#EF4444",
                "Medium Priority": "#F59E0B",
                "Low Priority":    "#22C55E",
            },
        )
        fig.update_layout(margin=dict(t=40, b=0, l=0, r=0), showlegend=False)
        col_r.plotly_chart(fig, use_container_width=True)

    col_l2, col_r2 = st.columns(2)

    # Unconfirmed Demand Amount by Key Account
    if account_df is not None and not account_df.empty:
        acct_col = account_df.columns[0]
        amt_cols = [c for c in account_df.columns if "unconfirmed demand" in c.lower()]
        if amt_cols:
            top_accts = account_df.sort_values(amt_cols[0], ascending=False).head(15)
            fig = px.bar(
                top_accts, x=acct_col, y=amt_cols[0],
                title="Unconfirmed Demand ($) by Key Account",
                labels={amt_cols[0]: "Unconfirmed Demand ($)", acct_col: "Key Account"},
                color=amt_cols[0],
                color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(coloraxis_showscale=False, margin=dict(t=40, b=0, l=0, r=0))
            col_l2.plotly_chart(fig, use_container_width=True)

    # Fill Rate by Key Account
    if account_df is not None and not account_df.empty:
        acct_col  = account_df.columns[0]
        rate_cols = [c for c in account_df.columns if "fill rate" in c.lower()]
        if rate_cols:
            fig = px.bar(
                account_df.sort_values(rate_cols[0]),
                x=rate_cols[0], y=acct_col, orientation="h",
                title="Fill Rate (%) by Key Account",
                labels={rate_cols[0]: "Fill Rate (%)", acct_col: "Key Account"},
                color=rate_cols[0],
                color_continuous_scale=["#EF4444", "#F59E0B", "#22C55E"],
                range_color=[0, 100],
            )
            fig.update_layout(
                coloraxis_showscale=False,
                margin=dict(t=40, b=0, l=0, r=0),
                yaxis=dict(autorange="reversed"),
            )
            col_r2.plotly_chart(fig, use_container_width=True)

    col_l3, col_r3 = st.columns(2)

    # Fill Rate by Plant
    if plant_df is not None and not plant_df.empty:
        plant_col = plant_df.columns[0]
        rate_cols = [c for c in plant_df.columns if "fill rate" in c.lower()]
        if rate_cols:
            fig = px.bar(
                plant_df.sort_values(rate_cols[0]),
                x=rate_cols[0], y=plant_col, orientation="h",
                title="Fill Rate (%) by Plant",
                labels={rate_cols[0]: "Fill Rate (%)", plant_col: "Plant"},
                color=rate_cols[0],
                color_continuous_scale=["#EF4444", "#F59E0B", "#22C55E"],
                range_color=[0, 100],
            )
            fig.update_layout(
                coloraxis_showscale=False,
                margin=dict(t=40, b=0, l=0, r=0),
                yaxis=dict(autorange="reversed"),
            )
            col_l3.plotly_chart(fig, use_container_width=True)

    # Unconfirmed Qty by Product (top 15)
    if product_df is not None and not product_df.empty:
        prod_col  = product_df.columns[0]
        qty_cols  = [c for c in product_df.columns if "unconfirmed qty" in c.lower()]
        if qty_cols:
            top15 = product_df.sort_values(qty_cols[0], ascending=False).head(15)
            fig = px.bar(
                top15, x=qty_cols[0], y=prod_col, orientation="h",
                title="Top 15 Products by Unconfirmed Qty",
                labels={qty_cols[0]: "Unconfirmed Qty", prod_col: "Product"},
                color=qty_cols[0],
                color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(
                coloraxis_showscale=False,
                margin=dict(t=40, b=0, l=0, r=0),
                yaxis=dict(autorange="reversed"),
            )
            col_r3.plotly_chart(fig, use_container_width=True)

# ── Unconfirmed Demand Report ─────────────────────────────────────────────────
with tab_unc:
    st.subheader(f"Unconfirmed Demand Report — {len(unconfirmed_df):,} problem lines")
    if unconfirmed_df.empty:
        st.success("No unconfirmed demand found in this dataset.")
    else:
        # Filters
        fc1, fc2, fc3, fc4, fc5 = st.columns(5)

        def _opts(col: str) -> list:
            if col not in unconfirmed_df.columns:
                return ["All"]
            return ["All"] + sorted(unconfirmed_df[col].dropna().unique().tolist())

        sel_acct  = fc1.selectbox("Key Account", _opts("key_account"),  key="unc_acct")
        sel_plant = fc2.selectbox("Plant",        _opts("plant"),        key="unc_plant")
        sel_prod  = fc3.selectbox("Product",      _opts("product"),      key="unc_prod")
        sel_cdm   = fc4.selectbox("CDM",          _opts("cdm_name"),     key="unc_cdm")
        sel_pri   = fc5.selectbox("Priority",     _opts("Priority"),     key="unc_pri")

        view = unconfirmed_df.copy()
        if sel_acct  != "All" and "key_account" in view.columns: view = view[view["key_account"]  == sel_acct]
        if sel_plant != "All" and "plant"        in view.columns: view = view[view["plant"]         == sel_plant]
        if sel_prod  != "All" and "product"      in view.columns: view = view[view["product"]       == sel_prod]
        if sel_cdm   != "All" and "cdm_name"     in view.columns: view = view[view["cdm_name"]      == sel_cdm]
        if sel_pri   != "All" and "Priority"     in view.columns: view = view[view["Priority"]      == sel_pri]

        st.caption(f"Showing {len(view):,} rows after filters.")
        st.dataframe(view, use_container_width=True, hide_index=True)

# ── Key Account Summary ───────────────────────────────────────────────────────
with tab_acct:
    st.subheader("Key Account Summary")
    if account_df is None or account_df.empty:
        st.info("Key Account column not found.")
    else:
        st.dataframe(account_df, use_container_width=True, hide_index=True)
        acct_col  = account_df.columns[0]
        amt_cols  = [c for c in account_df.columns if "unconfirmed demand" in c.lower()]
        if amt_cols:
            fig = px.bar(
                account_df.sort_values(amt_cols[0], ascending=False),
                x=acct_col, y=amt_cols[0],
                title="Unconfirmed Demand ($) by Key Account",
                color=amt_cols[0], color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ── Product Summary ───────────────────────────────────────────────────────────
with tab_prod:
    st.subheader("Product Summary")
    if product_df is None or product_df.empty:
        st.info("Product column not found.")
    else:
        st.dataframe(product_df, use_container_width=True, hide_index=True)
        prod_col = product_df.columns[0]
        amt_cols = [c for c in product_df.columns if "unconfirmed demand" in c.lower()]
        if amt_cols:
            top15 = product_df.sort_values(amt_cols[0], ascending=False).head(15)
            fig = px.bar(
                top15, x=prod_col, y=amt_cols[0],
                title="Top 15 Products by Unconfirmed Demand ($)",
                color=amt_cols[0], color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ── Plant Summary ─────────────────────────────────────────────────────────────
with tab_plant:
    st.subheader("Plant Summary")
    if plant_df is None or plant_df.empty:
        st.info("Plant column not found.")
    else:
        st.dataframe(plant_df, use_container_width=True, hide_index=True)
        plant_col = plant_df.columns[0]
        amt_cols  = [c for c in plant_df.columns if "unconfirmed demand" in c.lower()]
        if amt_cols:
            fig = px.bar(
                plant_df.sort_values(amt_cols[0], ascending=False),
                x=plant_col, y=amt_cols[0],
                title="Unconfirmed Demand ($) by Plant",
                color=amt_cols[0], color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ── CDM Summary ───────────────────────────────────────────────────────────────
with tab_cdm:
    st.subheader("CDM Summary")
    if cdm_df is None or cdm_df.empty:
        st.info("CDM Name column not found.")
    else:
        st.dataframe(cdm_df, use_container_width=True, hide_index=True)
        cdm_col  = cdm_df.columns[0]
        amt_cols = [c for c in cdm_df.columns if "unconfirmed demand" in c.lower()]
        if amt_cols:
            fig = px.bar(
                cdm_df.sort_values(amt_cols[0], ascending=False),
                x=cdm_col, y=amt_cols[0],
                title="Unconfirmed Demand ($) by CDM",
                color=amt_cols[0], color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ── Top 10 Issues ─────────────────────────────────────────────────────────────
with tab_top10:
    st.subheader("Top 10 Issues")
    sections = [
        ("Top 10 Key Accounts by Unconfirmed Demand", top10.get("accounts_by_unc_amount")),
        ("Top 10 Products by Unconfirmed Demand",     top10.get("products_by_unc_amount")),
        ("Top 10 Sales Orders by Unconfirmed Demand", top10.get("orders_by_unc_amount")),
        ("Top 10 CDMs by Unconfirmed Demand",         top10.get("cdm_by_unc_amount")),
    ]
    col_a, col_b = st.columns(2)
    for idx, (title, t_df) in enumerate(sections):
        col = col_a if idx % 2 == 0 else col_b
        col.markdown(f"**{title}**")
        if t_df is not None and not t_df.empty:
            col.dataframe(t_df, use_container_width=True, hide_index=True)
            val_col = t_df.columns[1]
            fig = px.bar(
                t_df, x=val_col, y=t_df.columns[0], orientation="h",
                color=val_col, color_continuous_scale=["#FEF3C7", "#EF4444"],
            )
            fig.update_layout(
                coloraxis_showscale=False, showlegend=False,
                margin=dict(t=10, b=0, l=0, r=0),
                yaxis=dict(autorange="reversed"),
                height=280,
            )
            col.plotly_chart(fig, use_container_width=True)
        else:
            col.info("No data available.")

# ── Raw Preview ───────────────────────────────────────────────────────────────
with tab_raw:
    st.subheader("Raw Export Preview")
    st.caption("First 500 rows as read from the file — before any cleaning.")
    st.dataframe(df_raw.head(500), use_container_width=True, hide_index=True)

    st.subheader("Cleaned Data Preview")
    st.caption("After header detection, type conversion, fill-down, and calculated columns.")
    st.dataframe(df_clean.head(500), use_container_width=True, hide_index=True)

# ── Download Report ───────────────────────────────────────────────────────────
with tab_dl:
    st.subheader("Download Excel Report")
    st.markdown(
        "Generates a fully formatted Excel workbook with 10 sheets: "
        "Instructions, Executive Summary, Clean Data, Unconfirmed Demand Report, "
        "Key Account Summary, Product Summary, Plant Summary, CDM Summary, "
        "Top 10 Issues, and Raw Export Preview."
    )

    if st.button("Generate Excel Report", type="primary"):
        with st.spinner("Building Excel workbook..."):
            excel_bytes = generate_excel_report(
                df_clean=df_clean,
                df_raw=df_raw,
                kpis=kpis,
                unconfirmed_df=unconfirmed_df,
                account_df=account_df if account_df is not None else pd.DataFrame(),
                product_df=product_df if product_df is not None else pd.DataFrame(),
                plant_df=plant_df   if plant_df   is not None else pd.DataFrame(),
                cdm_df=cdm_df       if cdm_df     is not None else pd.DataFrame(),
                top10=top10,
                threshold=threshold,
            )
        st.download_button(
            label="Download Sales_Order_Fill_Rate_Report.xlsx",
            data=excel_bytes,
            file_name="Sales_Order_Fill_Rate_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
