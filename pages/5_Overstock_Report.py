"""Overstock Report — flags excess stock approaching its last sellable date.

Upload the Materials inventory export and the Last Sell / BDM master, process,
preview, and download the finished three-sheet (Mississauga / Calgary / Surrey)
workbook built to the business's golden specification.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from src.overstock_engine import (
    REGION_PLANTS,
    OverstockError,
    build_overstock,
    default_window,
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

# Open date window — set both ends yourself. Defaults reproduce a same-day run
# but the result depends only on the dates below, never on today's weekday.
_def_floor, _def_cutoff = default_window(date.today())
st.markdown("**Date window**")
col_floor, col_cutoff = st.columns(2)
with col_floor:
    sled_floor = st.date_input(
        "Include shelf-life expiry on/after", value=_def_floor,
        help="Rows whose Shelf Life Expiration Date is earlier than this are dropped.",
    )
with col_cutoff:
    cutoff = st.date_input(
        "Last sell by date on/before", value=_def_cutoff,
        help="Rows whose last-sell-by date (SLED − Last Sell Day) is later than this are dropped.",
    )

if sled_floor > cutoff:
    st.warning(
        "The shelf-life floor is after the last-sell cutoff — that usually keeps "
        "nothing. Double-check the two dates."
    )

if materials_file is None or master_file is None:
    st.info("Upload both files to generate the report.")
    st.stop()

if not st.button("Process files", type="primary"):
    st.stop()


@st.cache_data(show_spinner=False)
def _process(mat_bytes: bytes, master_bytes: bytes, floor: date, cut: date):
    materials = load_materials(io.BytesIO(mat_bytes))
    master = load_master(io.BytesIO(master_bytes))
    sheets = build_overstock(
        materials, master, sled_floor=floor, last_sell_cutoff=cut,
    )
    return sheets


with st.spinner("Matching inventory to the master and applying overstock rules..."):
    try:
        sheets = _process(
            materials_file.getvalue(), master_file.getvalue(), sled_floor, cutoff,
        )
    except OverstockError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the files: {exc}")
        st.stop()

total_rows = sum(len(df) for df in sheets.values())
if total_rows == 0:
    st.warning("No overstock rows matched the rules for this date window.")
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
    file_name=f"Overstock report {date.today():%B %d %Y}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
