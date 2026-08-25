"""Supplier Summary Dashboard — weekly compliance check-in per supplier.

Reuses the same compliance engine as the Supplier Compliance Report (page 1)
to build one run's Supplier Summary, shows it as charts, and lets you save
that run under a snapshot date so week-over-week movement is trackable over
time — a small Postgres history table, same hosted DB pages 2/3 already use
for Column Variants. Nothing here changes Compliance % or bill-back; this is
a read-only view of the same numbers.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

from src.compliance_engine import build_report
from src.config import MONTH_NAMES
from src.portal_importer import PortalImportError, load_portal
from src.sap_importer import SapImportError, load_sap
from src.supplier_exceptions_ui import load_exceptions_or_empty
from src.supplier_summary_history import summary_to_snapshot_rows
from src.supplier_summary_history_ui import dsn, get_store

RATE_SCALE = ["#EF4444", "#F59E0B", "#22C55E"]  # red -> amber -> green, matches pages 2-4

st.title("Supplier Summary Dashboard")
st.caption(
    "Upload the same SAP + Portal exports as the Supplier Compliance Report "
    "(page 1) to see this month's Supplier Summary as charts, then save the "
    "run as a dated snapshot to track each supplier's compliance week over "
    "week. Uses the same compliance rules as page 1 — this page only reads "
    "and charts that data, it never changes Compliance % or bill-back."
)

col_sap, col_portal = st.columns(2)
with col_sap:
    sap_file = st.file_uploader("1. SAP Export (.xlsx)", type=["xlsx"], key="ssd_sap")
with col_portal:
    portal_file = st.file_uploader("2. Portal Export (.xlsx)", type=["xlsx"], key="ssd_portal")

today = date.today()
years = list(range(today.year - 3, today.year + 1))
months = list(range(1, 13))

col_year, col_month, col_snap = st.columns(3)
with col_year:
    sel_year = st.selectbox("Report Year", years, index=years.index(today.year))
with col_month:
    sel_month = st.selectbox(
        "Report Month", months, index=today.month - 1,
        format_func=lambda m: MONTH_NAMES[m - 1],
    )
with col_snap:
    snapshot_date = st.date_input(
        "Snapshot date", value=today,
        help=(
            "The date this check-in is saved under when you click Save "
            "below. Defaults to today — back-date it if you're catching up "
            "on a missed week."
        ),
    )

ready = sap_file is not None and portal_file is not None
if not ready:
    st.info("Upload both the SAP and Portal exports to generate the Supplier Summary.")
    st.stop()

if not st.button("Generate Supplier Summary", type="primary"):
    st.stop()

try:
    with st.spinner("Loading SAP file..."):
        sap_df = load_sap(sap_file)
except SapImportError as e:
    st.error(f"SAP file error: {e}")
    st.stop()

try:
    with st.spinner("Loading Portal file..."):
        portal_df = load_portal(portal_file, sel_year, sel_month)
except PortalImportError as e:
    st.error(f"Portal file error: {e}")
    st.stop()

exceptions, tracker_names, exceptions_error = load_exceptions_or_empty()
if exceptions_error:
    st.info(exceptions_error)

with st.spinner("Applying compliance rules..."):
    sheets = build_report(
        sap_df, portal_df, sel_year, sel_month,
        exceptions=exceptions, tracker_names=tracker_names,
    )

summary = sheets["Supplier Summary"]
if summary.empty:
    st.warning("No suppliers in scope for this month — nothing to chart.")
    st.stop()

# Numeric compliance %, computed the same way summary_to_snapshot_rows does
# (found / with-inbound), not parsed from the sheet's formatted "82.3%"
# string. A vendor with no inbound POs this month has nothing to be
# compliant about -- NaN, not 0%, so it's excluded from the chart and the
# average rather than dragging both down.
with_inbound = summary["SAP POs With Inbound Delivery"].astype(float)
found = summary["Portal Files Found"].astype(float)
compliance_num = (found / with_inbound.where(with_inbound > 0)) * 100
summary = summary.assign(**{"Compliance % (num)": compliance_num})

st.success("Supplier Summary generated.")

st.subheader("This run")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Suppliers In Scope", f"{len(summary):,}")
avg_pct = compliance_num.mean()
k2.metric("Avg Compliance %", f"{avg_pct:.1f}%" if pd.notna(avg_pct) else "n/a")
k3.metric("Total Missing Docs", f"{int(summary['Missing Portal Files'].sum()):,}")
k4.metric("Total Invalid Uploads", f"{int(summary['Invalid Portal Uploads'].sum()):,}")

rated = summary.dropna(subset=["Compliance % (num)"])
if not rated.empty:
    if len(rated) <= 5:
        top_n = len(rated)
    else:
        top_n = st.slider(
            "Show N lowest-compliance suppliers", min_value=5,
            max_value=len(rated), value=min(15, len(rated)),
        )
    worst = rated.sort_values("Compliance % (num)").head(top_n)
    fig = px.bar(
        worst, x="Compliance % (num)", y="Vendor Name", orientation="h",
        color="Compliance % (num)", color_continuous_scale=RATE_SCALE,
        range_color=[0, 100],
        text=worst["Compliance % (num)"].map(lambda v: f"{v:.0f}%"),
        category_orders={"Vendor Name": worst["Vendor Name"].tolist()},
        labels={"Compliance % (num)": "Compliance %", "Vendor Name": ""},
        title=f"Lowest {len(worst)} of {len(rated)} rated suppliers by Compliance %",
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(coloraxis_showscale=False, margin=dict(l=0))
    st.plotly_chart(fig, use_container_width=True)
else:
    st.caption("No supplier had any inbound POs this month — nothing to rate.")

st.subheader("Full Supplier Summary")
summary_display = summary.drop(columns=["Compliance % (num)"])
st.dataframe(summary_display, use_container_width=True, hide_index=True)

xlsx_buf = io.BytesIO()
with pd.ExcelWriter(xlsx_buf, engine="openpyxl") as writer:
    summary_display.to_excel(writer, sheet_name="Supplier Summary", index=False)
st.download_button(
    "⬇️ Download Supplier Summary (.xlsx)",
    data=xlsx_buf.getvalue(),
    file_name=f"Supplier_Summary_{sel_year}_{sel_month:02d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ---------------------------------------------------------------------------
# Save this run + week-over-week history (Postgres-backed, optional)
# ---------------------------------------------------------------------------
st.subheader("Save & track over time")

dsn_value = dsn()
store = None
if not dsn_value:
    st.caption(
        "ℹ️ Weekly history unavailable — database not configured. The "
        "charts above are still yours to read; they just won't be saved."
    )
else:
    try:
        store = get_store(dsn_value)
    except Exception as exc:  # noqa: BLE001 - never let the DB break the dashboard
        st.caption(
            f"⚠️ Weekly history unavailable ({type(exc).__name__}). The "
            "charts above are still yours to read; they just won't be saved."
        )

if store is not None:
    if st.button(f"💾 Save as the {snapshot_date:%b %d, %Y} snapshot"):
        rows = summary_to_snapshot_rows(summary, snapshot_date, sel_year, sel_month)
        try:
            n = store.save_snapshot(rows)
            st.success(
                f"Saved {n} supplier row(s) under {snapshot_date:%b %d, %Y}. "
                "Saving the same date again overwrites it, so re-running "
                "midday is safe."
            )
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not save snapshot: {type(exc).__name__}: {exc}")
            with st.expander("Full error detail"):
                st.exception(exc)

    try:
        snap_dates = store.list_snapshot_dates()
    except Exception as exc:  # noqa: BLE001
        snap_dates = []
        st.caption(f"⚠️ Could not load saved history ({type(exc).__name__}).")

    if len(snap_dates) < 2:
        st.caption(
            f"{len(snap_dates)} snapshot(s) saved so far. Save at least two "
            "weekly snapshots to see week-over-week movement."
        )
    else:
        newest, previous = store.latest_two_dates()
        st.markdown(f"**Week over week — {previous:%b %d} → {newest:%b %d}**")

        pair_hist = store.history(since=previous)
        pair_df = pd.DataFrame([h.__dict__ for h in pair_hist])
        pair_df = pair_df[pair_df["snapshot_date"].isin([newest, previous])]

        wide = pair_df.pivot_table(
            index=["vendor_number", "vendor_name"],
            columns="snapshot_date", values="compliance_pct",
        ).sort_index(axis=1)
        if len(wide.columns) < 2:
            st.caption(
                "The two most recent snapshots don't share any suppliers — "
                "nothing to compare yet."
            )
        else:
            prev_col, new_col = wide.columns[0], wide.columns[-1]
            movers = wide.reset_index()
            movers["Delta (pts)"] = (movers[new_col] - movers[prev_col]) * 100
            movers[f"{prev_col:%b %d}"] = movers[prev_col] * 100
            movers[f"{newest:%b %d}"] = movers[new_col] * 100
            movers["Trend"] = movers["Delta (pts)"].map(
                lambda d: "▲" if pd.notna(d) and d > 0.05
                else ("▼" if pd.notna(d) and d < -0.05 else "→")
            )
            movers = movers.drop(columns=[prev_col, new_col]).sort_values("Delta (pts)")
            movers = movers.rename(
                columns={"vendor_number": "Vendor Number", "vendor_name": "Vendor Name"}
            )
            movers.columns.name = None
            st.caption("Sorted worst movers first — biggest compliance drop at the top.")
            st.dataframe(
                movers, use_container_width=True, hide_index=True,
                column_config={
                    f"{prev_col:%b %d}": st.column_config.NumberColumn(format="%.1f%%"),
                    f"{newest:%b %d}": st.column_config.NumberColumn(format="%.1f%%"),
                    "Delta (pts)": st.column_config.NumberColumn(format="%+.1f"),
                },
            )

        st.markdown("**Compliance % over time, by supplier**")
        all_hist = store.history()
        all_hist_df = pd.DataFrame([h.__dict__ for h in all_hist])
        roster = (
            all_hist_df[["vendor_number", "vendor_name"]]
            .drop_duplicates()
            .sort_values("vendor_name")
        )
        options = {
            f"{r.vendor_name} ({r.vendor_number})": r.vendor_number
            for r in roster.itertuples(index=False)
        }
        picked = st.multiselect(
            "Suppliers to chart", list(options.keys()),
            default=list(options.keys())[:5],
            help="Defaults to the first 5 alphabetically — pick specific suppliers to compare.",
        )
        if picked:
            picked_numbers = [options[p] for p in picked]
            trend_df = all_hist_df[all_hist_df["vendor_number"].isin(picked_numbers)].copy()
            trend_df["Compliance %"] = trend_df["compliance_pct"] * 100
            fig = px.line(
                trend_df, x="snapshot_date", y="Compliance %", color="vendor_name",
                markers=True, labels={"snapshot_date": "Snapshot date", "vendor_name": "Supplier"},
            )
            fig.update_layout(yaxis_range=[0, 100])
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.caption("Pick at least one supplier to see its trend line.")
