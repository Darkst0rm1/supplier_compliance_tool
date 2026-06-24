"""Donate / Dispose List — stock at, near, or past expiry that can no longer be
sold (the mirror of the Overstock report).

Upload the same two exports as the Overstock page (Materials inventory + Last
Sell / BDM master), set the shelf-life cutoff, preview, and download the finished
three-sheet (Mississauga / Calgary / Surrey) workbook.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from src.donate_dispose_engine import (
    REGION_PLANTS,
    DonateDisposeError,
    build_donate_dispose,
    default_cutoff,
    generate_excel,
    load_master,
    load_materials,
)

st.title("Donate / Dispose List")
st.caption(
    "Upload the Materials inventory export and the Last Sell / BDM material "
    "master to list stock at, near, or past its shelf-life expiry (across every "
    "storage location) that should be donated or disposed of — one sheet per "
    "warehouse: Mississauga, Calgary, Surrey."
)

col_mat, col_master = st.columns(2)
with col_mat:
    materials_file = st.file_uploader(
        "1. Materials file (.xlsx)", type=["xlsx"], key="dd_materials",
    )
with col_master:
    master_file = st.file_uploader(
        "2. Last Sell / BDM Material Master (.xlsx)", type=["xlsx"], key="dd_master",
    )

# Open date window — set the cutoff yourself. The default reproduces a same-day
# run but the result depends only on the date below, never on today's weekday.
st.markdown("**Date window**")
sled_cutoff = st.date_input(
    "Include shelf-life expiry on/before", value=default_cutoff(date.today()),
    help="Rows whose Shelf Life Expiration Date is later than this are dropped "
         "(those are still sellable and belong on the Overstock report instead).",
)

if materials_file is None or master_file is None:
    st.info("Upload both files to generate the list.")
    st.stop()

if not st.button("Process files", type="primary"):
    st.stop()


@st.cache_data(show_spinner=False)
def _process(mat_bytes: bytes, master_bytes: bytes, cutoff: date):
    materials = load_materials(io.BytesIO(mat_bytes))
    master = load_master(io.BytesIO(master_bytes))
    return build_donate_dispose(materials, master, sled_cutoff=cutoff)


with st.spinner("Matching inventory to the master and applying donate/dispose rules..."):
    try:
        sheets = _process(materials_file.getvalue(), master_file.getvalue(), sled_cutoff)
    except DonateDisposeError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the files: {exc}")
        st.stop()

total_rows = sum(len(df) for df in sheets.values())
if total_rows == 0:
    st.warning("No donate/dispose rows matched the rules for this date window.")
    st.stop()

cols = st.columns(len(sheets))
for col, (name, df) in zip(cols, sheets.items()):
    col.metric(f"{name} ({', '.join(REGION_PLANTS[name])})", f"{len(df):,} rows")

st.subheader("Preview")
tabs = st.tabs(list(sheets.keys()))
for tab, (name, df) in zip(tabs, sheets.items()):
    with tab:
        if df.empty:
            st.info("No rows for this warehouse.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

st.subheader("Download")
with st.spinner("Building workbook..."):
    xlsx = generate_excel(sheets)
st.download_button(
    "⬇️ Download DonateDispose_List.xlsx",
    data=xlsx,
    file_name=f"DonateDispose list - {date.today():%B %d %Y}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
