"""Risky Inventory — remove the 90-day rows from the cumulative 180-day report.

The user downloads a 90-day report and a cumulative 180-day report (which
repeats the same first 90 days). This page automates the one manual step the
user does today: deleting from the 180-day file the rows already present in the
90-day file. Upload both files, process, view the results, and download the
finished four-sheet workbook (90D Detail / 90D Summary / 180D Detail /
180D Summary).
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from src.risky_inventory_engine import (
    RiskyInventoryError,
    generate_excel,
    load_detail,
    remove_duplicate_rows,
)

st.title("Risky Inventory")
st.caption(
    "Upload the 90-day report and the cumulative 180-day report. The 180-day "
    "report repeats the first 90 days; this removes those rows from it "
    "automatically (comparing complete inventory lines, not just Material), then "
    "builds the four-sheet workbook."
)

col90, col180 = st.columns(2)
with col90:
    file90 = st.file_uploader("1. 90-day report (.xlsx)", type=["xlsx"], key="ri_90")
with col180:
    file180 = st.file_uploader(
        "2. 180-day cumulative report (.xlsx)", type=["xlsx"], key="ri_180"
    )

if file90 is None or file180 is None:
    st.info("Upload both files to process.")
    st.stop()

if not st.button("Process files", type="primary"):
    st.stop()


@st.cache_data(show_spinner=False)
def _process(bytes90: bytes, bytes180: bytes):
    d90 = load_detail(io.BytesIO(bytes90))
    d180 = load_detail(io.BytesIO(bytes180))
    d180_clean = remove_duplicate_rows(d90, d180)
    xlsx = generate_excel(d90, d180_clean)
    return d90, d180, d180_clean, xlsx


with st.spinner("Removing the 90-day rows from the 180-day report..."):
    try:
        d90, d180, d180_clean, xlsx = _process(file90.getvalue(), file180.getvalue())
    except RiskyInventoryError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the files: {exc}")
        st.stop()

removed = len(d180) - len(d180_clean)
c1, c2, c3 = st.columns(3)
c1.metric("90-day rows", f"{len(d90):,}")
c2.metric("180-day rows (uploaded)", f"{len(d180):,}")
c3.metric("Removed (already in 90-day)", f"{removed:,}")
if removed == 0:
    st.info("No 90-day rows found in the 180-day file — it was already cleaned.")

st.subheader("Results")
tab90, tab180 = st.tabs(["90-day detail", f"Cleaned 180-day detail ({len(d180_clean):,} rows)"])
with tab90:
    st.dataframe(
        pd.DataFrame(d90.rows, columns=d90.headers),
        use_container_width=True, hide_index=True,
    )
with tab180:
    st.dataframe(
        pd.DataFrame(d180_clean.rows, columns=d180_clean.headers),
        use_container_width=True, hide_index=True,
    )

st.subheader("Download")
st.download_button(
    "Download Risky Inventory Report",
    data=xlsx,
    file_name=f"Risky Inventory Report - {date.today():%B %d %Y}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
