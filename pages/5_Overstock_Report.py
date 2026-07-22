"""Overstock Report — flags excess stock approaching its last sellable date.

Upload the Materials inventory export and the Last Sell / BDM master, optionally
add the per-plant EWM stock exports to name the bin each batch is sitting in,
then preview and download the finished three-sheet (Mississauga / Calgary /
Surrey) workbook built to the business's golden specification.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from src.overstock_engine import (
    EWM_PLANTS,
    OUT_BIN,
    REGION_PLANTS,
    OverstockError,
    build_overstock,
    default_window,
    generate_excel,
    load_ewm_bins,
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

# Optional bins. One export per plant — 2925 and 2935 have none, so their rows
# always come through with an empty Bin.
st.markdown("**3. EWM stock exports — optional, one per plant**")
st.caption(
    "Adds the Bin column (which bin each batch is sitting in). Skip any you "
    "don't have — the rest of the report is unaffected. Plants 2925 and 2935 "
    "have no EWM export, so those rows never get a bin."
)
ewm_files = {}
for col, plant in zip(st.columns(len(EWM_PLANTS)), EWM_PLANTS):
    with col:
        ewm_files[plant] = st.file_uploader(
            f"EWM {plant} (.xlsx)", type=["xlsx"], key=f"ovs_ewm_{plant}",
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
def _process(
    mat_bytes: bytes,
    master_bytes: bytes,
    floor: date,
    cut: date,
    ewm_bytes: tuple[tuple[str, bytes], ...],
):
    materials = load_materials(io.BytesIO(mat_bytes))
    master = load_master(io.BytesIO(master_bytes))
    bins = {
        plant: load_ewm_bins(io.BytesIO(raw), expected_plant=plant)
        for plant, raw in ewm_bytes
    }
    sheets = build_overstock(
        materials, master, sled_floor=floor, last_sell_cutoff=cut, ewm_bins=bins,
    )
    return sheets


# Tuple of pairs so the cache key covers exactly which plants were supplied.
ewm_bytes = tuple(
    (plant, f.getvalue()) for plant, f in ewm_files.items() if f is not None
)

with st.spinner("Matching inventory to the master and applying overstock rules..."):
    try:
        sheets = _process(
            materials_file.getvalue(), master_file.getvalue(), sled_floor, cutoff,
            ewm_bytes,
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

if ewm_bytes:
    # Unfilled bins are expected (2925/2935 have no export), so this is a
    # coverage note rather than a warning — but a sudden drop means the EWM
    # extract is a different vintage than the Materials snapshot.
    with_bin = sum(df[OUT_BIN].notna().sum() for df in sheets.values())
    supplied = ", ".join(p for p, _ in ewm_bytes)
    st.caption(
        f"Bins from EWM {supplied}: **{with_bin:,} of {total_rows:,}** rows "
        "matched. Rows with no bin are left blank — plants without an EWM "
        "export (2925, 2935) never match."
    )

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
