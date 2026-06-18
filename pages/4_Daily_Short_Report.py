"""Daily Short Report — order shortfall analysis built to the embedded template."""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src.daily_short_engine import (
    COL_CUSTOMER,
    COL_ORDER_TYPE,
    COL_PLANT,
    COL_REQ_DATE,
    COL_VENDOR_NAME,
    DailyShortError,
    SHORT_DEFS,
    build_group_summary,
    build_kpis,
    build_short_table,
    generate_excel_report,
    load_daily_short,
)

st.title("Daily Short Report")
st.caption(
    "Upload the SAPUI5 Daily Short export to see fill rates and the lines shorted "
    "at confirmation, outbound delivery, and invoicing — built to the template."
)

# ---------------------------------------------------------------------------
# Sidebar — upload + settings
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Upload & Settings")
    uploaded_file = st.file_uploader(
        "Daily Short Export (.xlsx)", type=["xlsx", "xls"], key="dsr_upload",
    )
    top_choice = st.selectbox(
        "Show top N per table", ["5", "10", "25", "50", "All"], index=0,
        help="The template specifies Top 5; change here to see more.",
    )
    top_n = None if top_choice == "All" else int(top_choice)
    st.caption(
        "**Note:** The embedded template/summary rows at the bottom of the export "
        "are detected and excluded automatically."
    )

if uploaded_file is None:
    st.info("Upload the Daily Short Export (.xlsx) using the sidebar to get started.")
    st.stop()


@st.cache_data(show_spinner=False)
def _process(file_bytes: bytes) -> pd.DataFrame:
    return load_daily_short(io.BytesIO(file_bytes))


with st.spinner("Reading and cleaning data..."):
    try:
        df_all = _process(uploaded_file.getvalue())
    except DailyShortError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read file: {exc}")
        st.stop()

if df_all.empty:
    st.warning("No order lines found in the file.")
    st.stop()

# ---------------------------------------------------------------------------
# Filters
# ---------------------------------------------------------------------------
if "dsr_filter_version" not in st.session_state:
    st.session_state["dsr_filter_version"] = 0
_v = st.session_state["dsr_filter_version"]


def _opts(col: str) -> list:
    if col not in df_all.columns:
        return []
    return sorted(df_all[col].dropna().astype(str).unique().tolist())


has_date = COL_REQ_DATE in df_all.columns and df_all[COL_REQ_DATE].notna().any()
if has_date:
    _dates = df_all[COL_REQ_DATE].dropna()
    min_date, max_date = _dates.min().date(), _dates.max().date()
else:
    min_date = max_date = date.today()

with st.sidebar:
    st.markdown("---")
    st.header("Filters")
    sel_plant = st.multiselect("Plant", _opts(COL_PLANT), key=f"dsr_plant_{_v}")
    sel_customer = st.multiselect("Customer", _opts(COL_CUSTOMER), key=f"dsr_cust_{_v}")
    sel_vendor = st.multiselect("Vendor", _opts(COL_VENDOR_NAME), key=f"dsr_vendor_{_v}")
    sel_type = st.multiselect("Sales Order Type", _opts(COL_ORDER_TYPE), key=f"dsr_type_{_v}")

    if has_date:
        date_range = st.date_input(
            "Requested Delivery Date",
            value=(min_date, max_date),
            min_value=min_date, max_value=max_date, key=f"dsr_date_{_v}",
        )
        if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
            d_start, d_end = date_range
        else:
            d_start = d_end = None
    else:
        d_start = d_end = None

    if st.button("Reset Filters", key="dsr_reset"):
        st.session_state["dsr_filter_version"] += 1
        st.rerun()

df = df_all.copy()
if sel_plant:
    df = df[df[COL_PLANT].isin(sel_plant)]
if sel_customer and COL_CUSTOMER in df.columns:
    df = df[df[COL_CUSTOMER].isin(sel_customer)]
if sel_vendor and COL_VENDOR_NAME in df.columns:
    df = df[df[COL_VENDOR_NAME].isin(sel_vendor)]
if sel_type and COL_ORDER_TYPE in df.columns:
    df = df[df[COL_ORDER_TYPE].isin(sel_type)]
if has_date and d_start and d_end:
    mask = (df[COL_REQ_DATE].dt.date >= d_start) & (df[COL_REQ_DATE].dt.date <= d_end)
    df = df[mask]
df = df.reset_index(drop=True)

if df.empty:
    st.warning("No lines match the current filters.")
    st.stop()

kpis = build_kpis(df)

_filtered = any([sel_plant, sel_customer, sel_vendor, sel_type,
                 (has_date and d_start and d_end and d_start != min_date)])
if _filtered:
    st.info(
        f"Filters active — **{len(df):,}** of **{len(df_all):,}** lines. "
        "Use **Reset Filters** in the sidebar to clear."
    )

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_over, tab_unc, tab_del, tab_inv, tab_dl = st.tabs([
    "Overview",
    "Unconfirmed",
    "Shorted at Delivery",
    "Shorted at Invoicing",
    "Download",
])

# ── Overview ─────────────────────────────────────────────────────────────────
with tab_over:
    st.subheader("Fill Rates")
    st.caption("Confirmation has one rate; outbound delivery and invoicing each have two.")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Confirmation (Conf/Ord)", f"{kpis['confirm_rate']:.2f}%")
    c2.metric("Delivery vs Ordered", f"{kpis['delivery_vs_order']:.2f}%")
    c3.metric("Delivery vs Confirmed", f"{kpis['delivery_vs_confirmed']:.2f}%")
    c4.metric("Invoice vs Ordered (Total)", f"{kpis['invoice_vs_order']:.2f}%")
    c5.metric("Invoice vs Delivered", f"{kpis['invoice_vs_delivered']:.2f}%")

    c6, c7, c8, c9, c10 = st.columns(5)
    c6.metric("Order Lines", f"{kpis['lines']:,}")
    c7.metric("Ordered", f"{kpis['ordered']:,.0f}")
    c8.metric("Confirmed", f"{kpis['confirmed']:,.0f}")
    c9.metric("Delivered", f"{kpis['delivered']:,.0f}")
    c10.metric("Invoiced", f"{kpis['invoiced']:,.0f}")

    st.markdown("---")
    col_l, col_r = st.columns(2)

    # Waterfall: start at Ordered, subtract the loss at each stage, end at Invoiced.
    ordered = kpis["ordered"] or 1  # guard divide-by-zero
    g1 = kpis["ordered"] - kpis["confirmed"]      # ordered not confirmed
    g2 = kpis["confirmed"] - kpis["delivered"]    # confirmed not delivered
    g3 = kpis["delivered"] - kpis["invoiced"]     # delivered not invoiced

    def _pof(v):  # percent of ordered
        return f"{v / ordered * 100:.1f}% of ord."

    fig_stage = go.Figure(go.Waterfall(
        orientation="v",
        measure=["absolute", "relative", "relative", "relative", "total"],
        x=["Ordered", "Not Confirmed", "Not Delivered", "Not Invoiced", "Invoiced"],
        y=[kpis["ordered"], -g1, -g2, -g3, kpis["invoiced"]],
        text=[
            f"{kpis['ordered']:,.0f}<br>100%",
            f"−{g1:,.0f}<br>{_pof(g1)}",
            f"−{g2:,.0f}<br>{_pof(g2)}",
            f"−{g3:,.0f}<br>{_pof(g3)}",
            f"{kpis['invoiced']:,.0f}<br>{kpis['invoice_vs_order']:.1f}% of ord.",
        ],
        textposition="outside",
        connector={"line": {"color": "#9AA0A6"}},
        decreasing={"marker": {"color": "#EF4444"}},
        increasing={"marker": {"color": "#22C55E"}},
        totals={"marker": {"color": "#1F4E79"}},
        hovertemplate="%{x}: %{y:,.0f}<extra></extra>",
    ))
    fig_stage.update_layout(
        title="Order → Invoice: where quantity is lost",
        margin=dict(t=50, b=0, l=0, r=0),
        yaxis_title="Cases (CS)",
        showlegend=False,
    )
    col_l.plotly_chart(fig_stage, use_container_width=True)

    short_df = pd.DataFrame({
        "Stage": ["Ordered − Confirmed", "Confirmed − Delivered", "Delivered − Invoiced"],
        "Short Qty": [kpis["short_unconfirmed"], kpis["short_delivery"], kpis["short_invoice"]],
        "Lines": [kpis["lines_unconfirmed"], kpis["lines_delivery"], kpis["lines_invoice"]],
    })
    fig_short = px.bar(
        short_df, x="Stage", y="Short Qty", title="Short Quantity by Stage (stage-to-stage gap)",
        text="Lines", color="Short Qty", color_continuous_scale=["#FEF3C7", "#EF4444"],
    )
    fig_short.update_traces(texttemplate="%{text} lines", textposition="outside")
    fig_short.update_layout(margin=dict(t=40, b=0, l=0, r=0), coloraxis_showscale=False)
    col_r.plotly_chart(fig_short, use_container_width=True)


def _render_short_tab(key: str, label: str, rates: list[tuple[str, float]], caption: str):
    scol = {"unconfirmed": "short_unconfirmed", "delivery": "short_delivery",
            "invoice": "short_invoice"}[key]
    lcol = {"unconfirmed": "lines_unconfirmed", "delivery": "lines_delivery",
            "invoice": "lines_invoice"}[key]

    cols = st.columns(len(rates) + 2)
    for col, (rlabel, rval) in zip(cols, rates):
        col.metric(rlabel, f"{rval:.2f}%")
    cols[-2].metric("Total Short (qty)", f"{kpis[scol]:,.0f}")
    cols[-1].metric("Lines Shorted", f"{kpis[lcol]:,}")

    table = build_short_table(df, key, top_n=top_n)
    heading = "all lines" if top_n is None else f"Top {top_n}"
    st.subheader(f"{label} — {heading}")
    st.caption(caption)
    if table.empty:
        st.success("No shorts found for this stage.")
        return

    st.dataframe(table, use_container_width=True, hide_index=True)

    fig = px.bar(
        table.head(top_n or 15),
        x="Shorted", y="Sales order #", orientation="h",
        hover_data=["Customer", "Material description"],
        color="Shorted", color_continuous_scale=["#FEF3C7", "#EF4444"],
        title=f"{label} — biggest shorts",
    )
    fig.update_layout(margin=dict(t=40, b=0, l=0, r=0), coloraxis_showscale=False,
                      yaxis=dict(autorange="reversed", type="category"))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("**Breakdowns**")
    bc1, bc2 = st.columns(2)
    cust = build_group_summary(df, COL_CUSTOMER, key)
    if cust is not None and not cust.empty:
        bc1.caption("Top customers by short qty")
        bc1.dataframe(cust.head(10), use_container_width=True, hide_index=True)
    vend = build_group_summary(df, COL_VENDOR_NAME, key)
    if vend is not None and not vend.empty:
        bc2.caption("Top vendors by short qty")
        bc2.dataframe(vend.head(10), use_container_width=True, hide_index=True)


with tab_unc:
    _render_short_tab(
        "unconfirmed", "Unconfirmed Quantities",
        rates=[("Confirmation (Conf/Ord)", kpis["confirm_rate"])],
        caption="Lines ordered but not fully confirmed. Shorted = Ordered − Confirmed.",
    )
with tab_del:
    _render_short_tab(
        "delivery", "Confirmed but No Outbound Delivery",
        rates=[("Delivery vs Ordered", kpis["delivery_vs_order"]),
               ("Delivery vs Confirmed", kpis["delivery_vs_confirmed"])],
        caption="Lines confirmed but not (fully) put on an outbound delivery. Shorted = Confirmed − Delivered.",
    )
with tab_inv:
    _render_short_tab(
        "invoice", "Delivered but Not Invoiced",
        rates=[("Invoice vs Ordered (Total)", kpis["invoice_vs_order"]),
               ("Invoice vs Delivered", kpis["invoice_vs_delivered"])],
        caption="Lines delivered but not (fully) invoiced. Shorted = Delivered − Invoiced.",
    )

# ── Download ─────────────────────────────────────────────────────────────────
with tab_dl:
    st.subheader("Download Excel Report")
    st.markdown(
        "A workbook with a **Summary** sheet (totals + fill rates) and one sheet "
        "per short analysis, using the template columns: Plant, Sales order #, "
        "Customer, Material, Material description, Ordered, Confirmed, Shorted, Reason."
    )
    if _filtered:
        st.info(f"The report reflects your active filters ({len(df):,} of {len(df_all):,} lines).")
    if st.button("Generate Excel Report", type="primary"):
        with st.spinner("Building workbook..."):
            xlsx = generate_excel_report(df, kpis, top_n=top_n)
        st.download_button(
            "⬇️ Download Daily_Short_Report.xlsx",
            data=xlsx,
            file_name="Daily_Short_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
