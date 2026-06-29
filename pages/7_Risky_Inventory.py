"""Risky Inventory — split one 0–180 day report into 0-90 / 91-180 buckets.

Upload the full 0–180 day report, pick the report run date, and the app buckets
each row by MRP Last Sell Date (cutoff = run date + 90 days), then builds a
workbook whose Summary sheet is a live Excel PivotTable filterable by Bucket.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from src.risky_inventory_engine import (
    BUCKET_0_90,
    BUCKET_91_180,
    BUCKET_NONE,
    RiskyInventoryError,
    assign_buckets,
    compute_cutoff,
    generate_excel,
    load_detail,
)

st.title("Risky Inventory")
st.caption(
    "Upload the full 0–180 day report. Rows are split by MRP Last Sell Date into "
    "0-90 and 91-180 day buckets (cutoff = report run date + 90 days). The "
    "downloaded Summary sheet is a live Excel PivotTable you can filter by Bucket."
)

uploaded = st.file_uploader("Risky Inventory 0–180 day report (.xlsx)", type=["xlsx"], key="ri_file")
run_date = st.date_input("Report run date", value=date.today(), key="ri_run_date")
cutoff = compute_cutoff(run_date)
st.caption(f"Cutoff: 0-90 = MRP Last Sell Date on/before **{cutoff:%b %d, %Y}**; later → 91-180.")

if uploaded is None:
    st.session_state.pop("ri_processed_file", None)
    st.info("Upload the 0–180 day report to begin.")
    st.stop()

if st.button("Process file", type="primary"):
    st.session_state["ri_processed_file"] = uploaded.file_id
if st.session_state.get("ri_processed_file") != uploaded.file_id:
    st.info("Click **Process file** to split the report.")
    st.stop()


@st.cache_data(show_spinner=False)
def _process(file_bytes: bytes, cutoff_iso: str):
    detail = load_detail(io.BytesIO(file_bytes))
    bucketed, counts = assign_buckets(detail, date.fromisoformat(cutoff_iso))
    xlsx = generate_excel(bucketed)
    return bucketed, counts, xlsx


with st.spinner("Splitting the report and building the workbook..."):
    try:
        bucketed, counts, xlsx = _process(uploaded.getvalue(), cutoff.isoformat())
    except RiskyInventoryError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the file: {exc}")
        st.stop()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total rows", f"{len(bucketed.rows):,}")
c2.metric("0-90 Day", f"{counts[BUCKET_0_90]:,}")
c3.metric("91-180 Day", f"{counts[BUCKET_91_180]:,}")
c4.metric("No Last Sell Date", f"{counts[BUCKET_NONE]:,}")

st.subheader("Detail")
st.dataframe(
    pd.DataFrame(bucketed.rows, columns=bucketed.headers),
    use_container_width=True, hide_index=True,
)

st.subheader("Download")
st.caption("The Summary sheet is a live PivotTable — filter by Bucket (and the other filters) in Excel.")
st.download_button(
    "Download Risky Inventory Report",
    data=xlsx,
    file_name=f"Risky Inventory Report - {date.today():%B %d %Y}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
