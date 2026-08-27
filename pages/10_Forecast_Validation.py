"""Forecast Validation — checks whether the current 2026/2027 sales forecast
makes sense against actual demand.

Upload the raw SAP Sales Order exports (any mix of the two column layouts —
one has Key Account #, the other has Plant / Plant Name) covering whatever
months you have, plus optionally a current-forecast file, and this builds the
Plant + Material table (matching the "output needed" template), a forecast
assessment, and supporting summaries. Order Quantity is the demand signal;
Invoice Quantity is shown alongside it as the fulfillment signal — an open
2026 order that hasn't invoiced yet is still real demand, not zero.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

from src.forecast_validation_engine import (
    MONTH_ABBR,
    REQUIRED_PLANTS,
    ForecastValidationError,
    build_data_quality,
    build_forecast_validation,
    build_item_detail,
    build_main_table,
    build_monthly_summary,
    build_plant_summary,
    combine_sales_orders,
    default_history_and_forecast_months,
    generate_excel,
    load_forecast,
    load_open_orders,
    load_open_po,
    load_sales_orders,
    load_stock_on_hand,
    resolve_open_order_plants,
)


def _display_label(c: str) -> str:
    """H|Sept / F|Sept / OO|Sept / PO|Sept all become bare 'Sept' if simply
    stripped, which collides into duplicate column names Streamlit/Arrow
    can't render sensibly -- keep them distinguishable instead."""
    prefixes = {"H|": "Hist ", "F|": "Fcst ", "OO|": "OpenOrd ", "PO|": "OpenPO "}
    for prefix, label in prefixes.items():
        if c.startswith(prefix):
            return f"{label}{c[len(prefix):].strip()}"
    return c


STATUS_COLORS = {
    "REASONABLE": "#22C55E", "REVIEW": "#F59E0B", "HIGH": "#EF4444",
    "LOW": "#EF4444", "NEW / NO HISTORY": "#6366F1", "NO AUGUST BASE": "#94A3B8",
    "PASTE FORECAST": "#94A3B8",
}

st.title("Forecast Validation")
st.caption(
    "Upload the raw SAP Sales Order exports (Aug-Dec 2025, 2026 months as you "
    "have them — any mix of the Key Account # or Plant column layout) and, "
    "optionally, your current forecast, Open Orders, Open PO, and Materials/"
    "stock files. Order Quantity is current demand; Invoice Quantity is what "
    "has actually invoiced — shown separately because open 2026 orders may "
    "not be invoiced yet."
)

col_so, col_fc = st.columns(2)
with col_so:
    so_files = st.file_uploader(
        "1. SAP Sales Order exports (.xlsx) — upload all months at once",
        type=["xlsx"], accept_multiple_files=True, key="fv_so",
    )
with col_fc:
    fc_file = st.file_uploader(
        "2. Current forecast (.xlsx) — optional", type=["xlsx"], key="fv_fc",
        help=(
            "Either the Plant/Mat#/monthly-columns layout, or a raw SAP MRP "
            "export (Material, Plnt, Del/finish, Planned qty). Skip this and "
            "the forecast columns show 'PASTE FORECAST' until you have one."
        ),
    )

st.markdown("**Optional — near-term supply/demand visibility**")
col_oo, col_po, col_stock = st.columns(3)
with col_oo:
    oo_file = st.file_uploader(
        "3. Open Orders (.xlsx) — optional", type=["xlsx"], key="fv_oo",
        help=(
            "Sales orders already booked with a future Requested Delivery "
            "Date (Sales Order, Material, Requested Delivery Date, Order "
            "Quantity). This export has no Plant column — Plant is resolved "
            "by matching Sales Order + Material back to the SAP Sales Order "
            "exports above; a row that can't be matched is excluded from "
            "plant-level Open Orders (see Data Quality)."
        ),
    )
with col_po:
    po_file = st.file_uploader(
        "4. Open PO (.xlsx) — optional", type=["xlsx"], key="fv_po",
        help="Inbound purchase orders not yet received (Plant, Material, Delivery Date, PO Quantity).",
    )
with col_stock:
    stock_file = st.file_uploader(
        "5. Materials / Stock on Hand (.xlsx) — optional", type=["xlsx"], key="fv_stock",
        help=(
            "The same Materials inventory export used on the Overstock and "
            "Risky Inventory pages (Plant, Material, Unrestricted Stock). "
            "Preferred, more complete Stock on Hand source — without it, "
            "Stock on Hand falls back to the incidental figure in the Open "
            "PO export, which only covers materials with an open PO."
        ),
    )

if not so_files:
    st.info("Upload at least one SAP Sales Order export to begin.")
    st.stop()

if not st.button("Process files", type="primary"):
    st.stop()


@st.cache_data(show_spinner=False)
def _process(so_bytes: tuple[bytes, ...], fc_bytes: bytes | None,
             oo_bytes: bytes | None, po_bytes: bytes | None, stock_bytes: bytes | None):
    dfs = [load_sales_orders(io.BytesIO(b)) for b in so_bytes]
    fact = combine_sales_orders(dfs)
    if fact.empty:
        raise ForecastValidationError("No usable rows found in the uploaded sales order files.")

    data_as_of = fact["Order Date"].max()
    hist_months, fc_months = default_history_and_forecast_months(data_as_of.date())

    if fc_bytes is not None:
        forecast = load_forecast(io.BytesIO(fc_bytes), fc_months)
    else:
        forecast = pd.DataFrame(columns=["Plant", "Material", "Forecast Year", "Forecast Month",
                                          "Forecast Qty", "Buyer Name", "Brand Name", "Material Description"])

    if oo_bytes is not None:
        open_orders = resolve_open_order_plants(load_open_orders(io.BytesIO(oo_bytes)), fact)
    else:
        open_orders = pd.DataFrame(columns=["Sales Order", "Material", "Plant", "Year", "Month", "Open Order Qty"])

    open_po = load_open_po(io.BytesIO(po_bytes)) if po_bytes is not None else \
        pd.DataFrame(columns=["Plant", "Material", "Year", "Month", "Open PO Qty", "Stock on Hand"])

    stock_on_hand = load_stock_on_hand(io.BytesIO(stock_bytes)) if stock_bytes is not None else \
        pd.DataFrame(columns=["Plant", "Material", "Stock on Hand"])

    main_table = build_main_table(fact, forecast, hist_months, fc_months,
                                   open_orders=open_orders, open_po=open_po, stock_on_hand=stock_on_hand)
    validation = build_forecast_validation(fact, forecast, fc_months, data_as_of)
    plant_summary = build_plant_summary(fact, data_as_of)
    item_detail = build_item_detail(fact)
    monthly_summary = build_monthly_summary(fact, data_as_of)
    data_quality = build_data_quality(fact, forecast, hist_months,
                                       open_orders=open_orders, open_po=open_po, stock_on_hand=stock_on_hand)

    xlsx = generate_excel(main_table, validation, plant_summary, item_detail,
                           monthly_summary, data_quality, hist_months, fc_months, data_as_of)

    return {
        "fact": fact, "forecast": forecast, "main_table": main_table,
        "validation": validation, "plant_summary": plant_summary,
        "item_detail": item_detail, "monthly_summary": monthly_summary,
        "data_quality": data_quality, "hist_months": hist_months,
        "fc_months": fc_months, "data_as_of": data_as_of, "xlsx": xlsx,
        "open_orders": open_orders, "open_po": open_po, "stock_on_hand": stock_on_hand,
    }


with st.spinner("Combining sales order exports and validating the forecast..."):
    try:
        result = _process(
            tuple(f.getvalue() for f in so_files),
            fc_file.getvalue() if fc_file is not None else None,
            oo_file.getvalue() if oo_file is not None else None,
            po_file.getvalue() if po_file is not None else None,
            stock_file.getvalue() if stock_file is not None else None,
        )
    except ForecastValidationError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the files: {exc}")
        st.stop()

main_table = result["main_table"]
validation = result["validation"]
plant_summary = result["plant_summary"]
item_detail = result["item_detail"]
monthly_summary = result["monthly_summary"]
data_quality = result["data_quality"]
data_as_of = result["data_as_of"]

st.success(f"Processed {len(result['fact']):,} sales order lines. Data as of {data_as_of:%b %d, %Y}.")
if fc_file is None:
    st.info(
        "No forecast file uploaded — the Forecast columns and Assessment "
        "show 'PASTE FORECAST' until one is provided."
    )
if oo_file is None and po_file is None:
    st.caption(
        "Open Orders / Open PO not uploaded — those columns and Stock on "
        "Hand are blank in the table and download below."
    )

# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
total_forecast = validation["Current Forecast"].dropna().sum() if not validation.empty else 0
company_order_growth = plant_summary["Order Growth %"].mean(skipna=True) if not plant_summary.empty else None
invoice_completion = (
    monthly_summary.loc[monthly_summary["Period"] == f"{data_as_of:%b %Y}", "Invoice Completion %"].iloc[0]
    if not monthly_summary.empty and (monthly_summary["Period"] == f"{data_as_of:%b %Y}").any() else None
)
status_counts = validation["Assessment"].value_counts() if not validation.empty else pd.Series(dtype=int)
oo_cols = [c for c in main_table.columns if c.startswith("OO|")]
po_cols = [c for c in main_table.columns if c.startswith("PO|")]
total_open_orders = main_table[oo_cols].sum().sum() if oo_cols else 0
total_open_po = main_table[po_cols].sum().sum() if po_cols else 0
total_stock = main_table["Stock on Hand"].sum() if "Stock on Hand" in main_table.columns else 0

st.subheader("Dashboard")
k1, k2, k3 = st.columns(3)
k1.metric("Total Forecast Cases", f"{total_forecast:,.0f}")
k2.metric("Avg Plant Order Growth %", f"{company_order_growth * 100:,.1f}%" if pd.notna(company_order_growth) else "n/a")
k3.metric("Invoice Completion % (current month)", f"{invoice_completion * 100:,.1f}%" if pd.notna(invoice_completion) else "n/a")

k1b, k2b, k3b = st.columns(3)
k1b.metric("Open Orders (next 4 months)", f"{total_open_orders:,.0f}")
k2b.metric("Open PO (next 4 months)", f"{total_open_po:,.0f}")
k3b.metric("Stock on Hand", f"{total_stock:,.0f}")

k4, k5, k6, k7 = st.columns(4)
k4.metric("Reasonable", int(status_counts.get("REASONABLE", 0)))
k5.metric("Review", int(status_counts.get("REVIEW", 0)))
k6.metric("High", int(status_counts.get("HIGH", 0)))
k7.metric("Low", int(status_counts.get("LOW", 0)))

# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
st.subheader("Charts")
chart_tabs = st.tabs(["Actual vs Forecast", "Order vs Invoice", "Forecast Status"])

with chart_tabs[0]:
    # Month labels ("Sept", "Oct"...) repeat identically between the history
    # and forecast halves of the main table -- a categorical x-axis keyed on
    # the bare label would collapse "Sept 2025" and "Sept 2026" onto the same
    # tick. Build "Mon YYYY" labels from the actual (year, month) pairs so
    # all 24 months plot in true sequence.
    hist_cols = [c for c in main_table.columns if c.startswith("H|")]
    fc_cols = [c for c in main_table.columns if c.startswith("F|")]
    hist_labels = [f"{MONTH_ABBR[m - 1]} {y}" for y, m in result["hist_months"]]
    fc_labels = [f"{MONTH_ABBR[m - 1]} {y}" for y, m in result["fc_months"]]

    hist_totals = pd.DataFrame({
        "Month": hist_labels, "Qty": main_table[hist_cols].sum().to_numpy(),
        "Series": "Actual Invoiced", "Order": range(len(hist_cols)),
    })
    fc_totals = pd.DataFrame({
        "Month": fc_labels, "Qty": main_table[fc_cols].sum().to_numpy(),
        "Series": "Forecast", "Order": range(len(hist_cols), len(hist_cols) + len(fc_cols)),
    })
    combined_chart = pd.concat([hist_totals, fc_totals])
    fig = px.line(
        combined_chart.sort_values("Order"), x="Month", y="Qty", color="Series", markers=True,
        category_orders={"Month": hist_labels + fc_labels},
    )
    st.plotly_chart(fig, use_container_width=True)

with chart_tabs[1]:
    fig2 = px.bar(
        monthly_summary, x="Period", y=["Order Qty", "Invoice Qty"], barmode="group",
        labels={"value": "Cases", "variable": ""},
    )
    st.plotly_chart(fig2, use_container_width=True)

with chart_tabs[2]:
    if status_counts.empty:
        st.caption("No forecast rows to assess yet.")
    else:
        sc_df = status_counts.reset_index()
        sc_df.columns = ["Status", "Count"]
        fig3 = px.bar(
            sc_df, x="Status", y="Count", color="Status",
            color_discrete_map=STATUS_COLORS,
        )
        fig3.update_layout(showlegend=False)
        st.plotly_chart(fig3, use_container_width=True)

# ---------------------------------------------------------------------------
# Filters
# ---------------------------------------------------------------------------
st.subheader("Main table")

f1, f2, f3, f4 = st.columns(4)
with f1:
    plant_filter = st.multiselect("Plant", REQUIRED_PLANTS, default=[])
with f2:
    brand_filter = st.multiselect("Brand", sorted([b for b in main_table["Brand name"].unique() if b]), default=[])
with f3:
    buyer_filter = st.multiselect("Buyer", sorted([b for b in main_table["Buyer name"].unique() if b]), default=[])
with f4:
    status_filter = st.multiselect("Forecast Status", sorted(validation["Assessment"].unique()) if not validation.empty else [], default=[])

search = st.text_input("Search Material # or Description")

filtered = main_table.copy()
if plant_filter:
    filtered = filtered[filtered["Plant"].astype(str).isin(plant_filter)]
if brand_filter:
    filtered = filtered[filtered["Brand name"].isin(brand_filter)]
if buyer_filter:
    filtered = filtered[filtered["Buyer name"].isin(buyer_filter)]
if search:
    s = search.strip().lower()
    filtered = filtered[
        filtered["Mat #"].str.lower().str.contains(s, na=False)
        | filtered["Desc"].str.lower().str.contains(s, na=False)
    ]
if status_filter and not validation.empty:
    flagged_keys = set(zip(
        validation.loc[validation["Assessment"].isin(status_filter), "Plant"],
        validation.loc[validation["Assessment"].isin(status_filter), "Material"],
    ))
    filtered = filtered[filtered.apply(lambda r: (str(r["Plant"]), r["Mat #"]) in flagged_keys, axis=1)]

display = filtered.rename(columns=_display_label)
st.dataframe(display, use_container_width=True, hide_index=True)
st.caption(f"{len(filtered):,} of {len(main_table):,} Plant + Material rows shown.")

# ---------------------------------------------------------------------------
# Material Detail
# ---------------------------------------------------------------------------
st.subheader("Material Detail")
mat_options = sorted(main_table["Mat #"].unique())
if mat_options:
    picked_mat = st.selectbox("Select a Material", mat_options)
    detail_rows = main_table[main_table["Mat #"] == picked_mat]
    st.dataframe(detail_rows.rename(columns=_display_label),
                 use_container_width=True, hide_index=True)

    detail_validation = validation[validation["Material"] == picked_mat] if not validation.empty else pd.DataFrame()
    if not detail_validation.empty:
        st.markdown("**Forecast assessment for this material**")
        st.dataframe(detail_validation, use_container_width=True, hide_index=True)

    detail_history = item_detail[item_detail["Material"] == picked_mat]
    if not detail_history.empty:
        fig4 = px.line(detail_history, x="Period", y=["Order Qty", "Invoice Qty"], color="Plant",
                        markers=True, labels={"value": "Cases", "variable": ""})
        st.plotly_chart(fig4, use_container_width=True)
else:
    st.caption("No materials to show yet.")

# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------
st.subheader("Download")
st.download_button(
    "⬇️ Download Forecast Excel",
    data=result["xlsx"],
    file_name=f"Forecast_Validation_{date.today():%Y-%m-%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Data Quality"):
    st.dataframe(data_quality, use_container_width=True, hide_index=True)
