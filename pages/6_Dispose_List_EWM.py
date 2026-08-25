"""Dispose List (EWM) — the warehouse's bin-level view of stock to dispose of.

Upload the per-plant EWM dispose exports and the Last Sell / BDM master,
process, preview, and download the finished workbook (one sheet per plant:
2910, 2920, 2930, plus 2925/2935 if that materials export is supplied) built
to the business's golden specification.
"""
from __future__ import annotations

import io
from datetime import date

import streamlit as st

from src.dispose_ewm_engine import (
    EXCLUDED_PRODUCT_PREFIXES,
    MATERIALS_PLANTS,
    OUT_BDM,
    PLANTS,
    SLED_CUTOFF_OFFSET_DAYS,
    DisposeEwmError,
    build_dispose_ewm,
    generate_excel,
    load_ewm,
    load_master,
    load_materials_dispose,
)

st.title("Dispose List (EWM)")
st.caption(
    "Upload the per-plant EWM dispose exports and the Last Sell / BDM material "
    "master. The export is kept as-is — every column, original row order — with "
    "the Brand Manager added and packaging materials removed. One sheet per "
    "plant: 2910, 2920, 2930, plus 2925 and 2935 if that materials export is "
    "supplied."
)

st.markdown("**1. EWM dispose exports — one per plant**")
st.caption("Upload at least one. Only the plants you supply get a sheet.")
ewm_files = {}
for col, plant in zip(st.columns(len(PLANTS)), PLANTS):
    with col:
        ewm_files[plant] = st.file_uploader(
            f"EWM {plant} dispose (.xlsx)", type=["xlsx"], key=f"dewm_{plant}",
        )

st.markdown("**2. Last Sell / BDM Material Master**")
master_file = st.file_uploader(
    "Material master (.xlsx)", type=["xlsx"], key="dewm_master",
    help="Used only to look up the Brand Manager for each product.",
)

st.markdown(f"**3. Materials dispose export — {' / '.join(MATERIALS_PLANTS)} (optional)**")
materials_file = st.file_uploader(
    "Materials dispose export (.xlsx)", type=["xlsx"], key="dewm_materials",
    help=(
        f"Plants {' and '.join(MATERIALS_PLANTS)} have no EWM system, so they "
        "never get an EWM export. Upload the Materials-based dispose extract "
        "instead — same shape the Donate/Dispose report produces, BDM already "
        "filled in. One file can carry both plants; it's split into two sheets."
    ),
)

st.caption(
    "Rows whose Product/Material number starts with "
    f"{' or '.join(EXCLUDED_PRODUCT_PREFIXES)} are excluded — packaging, "
    "display, label and sample material, never sellable stock."
)

materials_cutoff = None
if materials_file is not None:
    materials_cutoff = st.date_input(
        "Dispose cutoff date (2925/2935 only)",
        value=date.today(),
        help=(
            "A 2925/2935 row is kept only if both Shelf Life Expiration Date "
            f"and Last sell date fall on/before this date + "
            f"{SLED_CUTOFF_OFFSET_DAYS} days. Rows missing either date are "
            "dropped. Last sell date is computed from the master's Last Sell "
            "Day when the export doesn't already carry it."
        ),
    )
    if master_file is None:
        st.warning(
            "No material master uploaded — Last sell date can't be computed "
            "for a raw materials export, so this filter will drop 2925/2935 "
            "rows that don't already carry Last sell date."
        )

supplied = {p: f for p, f in ewm_files.items() if f is not None}
if (not supplied or master_file is None) and materials_file is None:
    st.info(
        "Upload at least one EWM dispose export and the material master, or "
        "the 2925/2935 materials dispose export, to generate the list."
    )
    st.stop()

if not st.button("Process files", type="primary"):
    st.stop()


@st.cache_data(show_spinner=False)
def _process(
    ewm_bytes: tuple[tuple[str, bytes], ...],
    master_bytes: bytes | None,
    materials_bytes: bytes | None,
    materials_cutoff: date | None,
):
    ewm = {
        plant: load_ewm(io.BytesIO(raw), expected_plant=plant)
        for plant, raw in ewm_bytes
    }
    bdm_lut = load_master(io.BytesIO(master_bytes)) if master_bytes else None
    materials_dispose = (
        load_materials_dispose(io.BytesIO(materials_bytes))
        if materials_bytes else None
    )
    return build_dispose_ewm(
        ewm, bdm_lut, materials_dispose, materials_cutoff=materials_cutoff
    )


ewm_bytes = tuple((p, f.getvalue()) for p, f in supplied.items())

with st.spinner("Applying the packaging exclusion and looking up brand managers..."):
    try:
        sheets = _process(
            ewm_bytes,
            master_file.getvalue() if master_file else None,
            materials_file.getvalue() if materials_file else None,
            materials_cutoff,
        )
    except DisposeEwmError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the files: {exc}")
        st.stop()

total_rows = sum(len(df) for df in sheets.values())
if total_rows == 0:
    st.warning("Every row was excluded — check you uploaded the dispose exports.")
    st.stop()

cols = st.columns(len(sheets))
for col, (plant, df) in zip(cols, sheets.items()):
    col.metric(f"Plant {plant}", f"{len(df):,} rows")

# A product missing from the master gets a blank BDM. Worth surfacing: it
# usually means the master is a different vintage than the EWM extract.
missing_bdm = sum(
    int((df[OUT_BDM].fillna("").astype(str).str.strip() == "").sum())
    for df in sheets.values()
)
if missing_bdm:
    st.caption(
        f"{missing_bdm:,} of {total_rows:,} rows have no Brand Manager — those "
        "products aren't in the master you uploaded."
    )

st.subheader("Preview")
tabs = st.tabs([f"Plant {p}" for p in sheets])
for tab, (plant, df) in zip(tabs, sheets.items()):
    with tab:
        if df.empty:
            st.info("No rows for this plant.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

st.subheader("Download")
with st.spinner("Building workbook..."):
    xlsx = generate_excel(sheets)
st.download_button(
    "⬇️ Download Dispose_List_EWM.xlsx",
    data=xlsx,
    file_name=f"Dispose list {date.today():%m%d} EWM.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
