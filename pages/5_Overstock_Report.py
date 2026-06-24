"""Overstock Report — flags excess stock approaching its last sellable date.

Upload the Materials inventory export and the Last Sell / BDM master, process,
preview, and download the finished three-sheet (Mississauga / Calgary / Surrey)
workbook built to the business's golden specification.
"""
from __future__ import annotations

import io
from datetime import date, timedelta

import pandas as pd
import streamlit as st

from src.overstock_engine import (
    LAST_SELL_CUTOFF_OFFSET_DAYS,
    REGION_PLANTS,
    SLED_FLOOR_OFFSET_DAYS,
    OverstockError,
    build_overstock,
    generate_excel,
    load_master,
    load_materials,
)

st.title("Overstock Report")
st.caption(
    "Upload the Materials inventory export and the Last Sell / BDM material "
    "master, then download the weekly overstock workbook (one sheet per "
    "warehouse: Mississauga, Calgary, Surrey)."
)

col_mat, col_master = st.columns(2)
with col_mat:
    materials_file = st.file_uploader(
        "1. Materials file (.xlsx)", type=["xlsx"], key="ovs_materials",
    )
with col_master:
    master_file = st.file_uploader(
        "2. Last Sell / BDM Material Master (.xlsx)", type=["xlsx"], key="ovs_master",
    )

report_dt = st.date_input(
    "Report date", value=date.today(),
    help="The run date. Items are included when their Shelf Life Expiration "
         f"Date is on/after this date + {SLED_FLOOR_OFFSET_DAYS} days and their "
         f"last-sell-by date is on/before this date + {LAST_SELL_CUTOFF_OFFSET_DAYS} days.",
)
sled_floor = report_dt + timedelta(days=SLED_FLOOR_OFFSET_DAYS)
cutoff = report_dt + timedelta(days=LAST_SELL_CUTOFF_OFFSET_DAYS)
st.caption(
    f"Window for this run — Shelf Life Expiration on/after **{sled_floor:%m/%d/%Y}**, "
    f"last sell by date on/before **{cutoff:%m/%d/%Y}**."
)

if materials_file is None or master_file is None:
    st.info("Upload both files to generate the report.")
    st.stop()

if not st.button("Process files", type="primary"):
    st.stop()


@st.cache_data(show_spinner=False)
def _process(mat_bytes: bytes, master_bytes: bytes, rpt: date):
    materials = load_materials(io.BytesIO(mat_bytes))
    master = load_master(io.BytesIO(master_bytes))
    sheets = build_overstock(materials, master, report_date=rpt)
    return sheets


with st.spinner("Matching inventory to the master and applying overstock rules..."):
    try:
        sheets = _process(materials_file.getvalue(), master_file.getvalue(), report_dt)
    except OverstockError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the files: {exc}")
        st.stop()

total_rows = sum(len(df) for df in sheets.values())
if total_rows == 0:
    st.warning("No overstock rows matched the rules for this report date.")
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
    "⬇️ Download Overstock_Report.xlsx",
    data=xlsx,
    file_name=f"Overstock report {report_dt:%B %d %Y}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
