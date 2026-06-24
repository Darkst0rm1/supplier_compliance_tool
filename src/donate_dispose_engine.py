"""Processing engine for the weekly Donate / Dispose list.

This is the mirror image of the Overstock report (see ``overstock_engine``).
Both read the same two SAPUI5 exports (Materials inventory + Last Sell / BDM
master) and both split into Mississauga / Calgary / Surrey sheets, but the date
window is flipped:

* **Overstock** keeps stock whose Shelf Life Expiration Date is *far enough out*
  to still be sold (and restricts to the main warehouse).
* **Donate / Dispose** keeps stock that is *at, near, or past* expiry — too late
  to sell, so it must be donated or disposed of — across *every* storage
  location (main warehouse, overstock, clearance, rework hold, consignment).

The supplied finished workbook ``DonateDispose list - June 24 2026 P2.xlsx`` is
the golden specification. Rules reverse-engineered against it (all three sheets
reproduced exactly — 48 / 92 / 19 rows):

1. Total stock = Unrestricted + Quality Inspection + Blocked > 0.
2. Plant belongs to the sheet's region (shared ``REGION_PLANTS``).
3. Material number does NOT start with "40" (display / shipper / label / sample
   packaging, never sellable stock) — shared with overstock.
4. The Material matches a master ``Product Number`` (so it has a Last Sell Day).
5. NOT the RANA retail brand handled by Sandra; NOT the Sweet Street ("SSD …")
   brand — shared with overstock.
6. Shelf Life Expiration Date present and on/before the cutoff
   (report date + ``SLED_CUTOFF_OFFSET_DAYS``).

There is **no** storage-location restriction and **no** last-sell-date filter
(the SLED cutoff alone defines the window). Rows are sorted by Shelf Life
Expiration Date ascending within each sheet.
"""
from __future__ import annotations

import io
from datetime import date, datetime, timedelta
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.overstock_engine import (
    EXCLUDED_MATERIAL_PREFIXES,
    MASTER_BDM_NAME,
    MASTER_LAST_SELL,
    MASTER_PRODUCT,
    MAT_BATCH,
    MAT_BLOCKED,
    MAT_DESCRIPTION,
    MAT_MATERIAL,
    MAT_PLANT,
    MAT_PLANT_NAME,
    MAT_QUALITY,
    MAT_SLED,
    MAT_SPECIAL_STOCK,
    MAT_STORAGE_DESC,
    MAT_STORAGE_LOC,
    MAT_TEXT_COLS,
    MAT_UNRESTRICTED,
    OUT_BDM,
    RANA_DESC_PREFIX,
    RANA_EXCLUDED_BDM,
    REGION_PLANTS,
    STOCK_COLS,
    SWEET_STREET_DESC_PREFIX,
    OverstockError,
    _coerce_report_date,
    load_master,
    load_materials,
)

# Re-export the shared loaders/error so the page imports from one module.
DonateDisposeError = OverstockError
__all__ = [
    "DonateDisposeError",
    "REGION_PLANTS",
    "SLED_CUTOFF_OFFSET_DAYS",
    "build_donate_dispose",
    "default_cutoff",
    "generate_excel",
    "load_master",
    "load_materials",
]

# ---------------------------------------------------------------------------
# Output columns (note: this report uses "Last sell day" / "Last sell date",
# without the "by" the overstock report uses).
# ---------------------------------------------------------------------------
OUT_LAST_SELL_DAY = "Last sell day"
OUT_LAST_SELL_DT = "Last sell date"

OUTPUT_COLUMNS = [
    MAT_MATERIAL,
    MAT_DESCRIPTION,
    MAT_PLANT,
    MAT_PLANT_NAME,
    OUT_BDM,
    MAT_STORAGE_LOC,
    MAT_STORAGE_DESC,
    MAT_BATCH,
    MAT_SLED,
    OUT_LAST_SELL_DAY,
    OUT_LAST_SELL_DT,
    MAT_SPECIAL_STOCK,
    MAT_UNRESTRICTED,
    MAT_QUALITY,
    MAT_BLOCKED,
]

# Date window, relative to the report run date. Stock is in scope when its
# Shelf Life Expiration Date is on/before report_date + this many days.
SLED_CUTOFF_OFFSET_DAYS = 4


def default_cutoff(report_date: date | datetime | None = None) -> date:
    """The default SLED cutoff for a run date. Used to pre-fill the page; the
    user can override it freely."""
    return _coerce_report_date(report_date) + timedelta(days=SLED_CUTOFF_OFFSET_DAYS)


def build_donate_dispose(
    materials: pd.DataFrame,
    master: pd.DataFrame,
    sled_cutoff: date | datetime | None = None,
    report_date: date | datetime | None = None,
) -> dict[str, pd.DataFrame]:
    """Apply the donate/dispose selection rules and return one DataFrame per
    region in finished-workbook column order.

    Pass ``sled_cutoff`` (include Shelf Life Expiration on/before) explicitly;
    left as ``None`` it falls back to report_date + ``SLED_CUTOFF_OFFSET_DAYS``
    (today if ``report_date`` is also ``None``)."""
    cutoff = pd.Timestamp(
        sled_cutoff if sled_cutoff is not None else default_cutoff(report_date)
    )

    m = materials.copy()
    lut = master.rename(columns={
        MASTER_BDM_NAME: OUT_BDM,
        MASTER_LAST_SELL: OUT_LAST_SELL_DAY,
    })
    m = m.merge(lut, how="left", left_on=MAT_MATERIAL, right_on=MASTER_PRODUCT)

    m[OUT_LAST_SELL_DT] = m[MAT_SLED] - pd.to_timedelta(m[OUT_LAST_SELL_DAY], unit="D")
    total_stock = m[STOCK_COLS].sum(axis=1)

    desc = m[MAT_DESCRIPTION].astype(str).str.strip().str.upper()
    bdm = m[OUT_BDM].astype(str).str.strip().str.upper()

    is_packaging = m[MAT_MATERIAL].astype(str).str.startswith(EXCLUDED_MATERIAL_PREFIXES)
    is_sweet_street = desc.str.startswith(SWEET_STREET_DESC_PREFIX)
    is_rana_sandra = (bdm == RANA_EXCLUDED_BDM) & desc.str.startswith(RANA_DESC_PREFIX)

    keep = (
        (total_stock > 0)
        & m[MAT_SLED].notna()
        & m[OUT_LAST_SELL_DAY].notna()
        & ~is_packaging
        & ~is_sweet_street
        & ~is_rana_sandra
        & (m[MAT_SLED] <= cutoff)
    )
    selected = m[keep].copy()

    sheets: dict[str, pd.DataFrame] = {}
    for sheet_name, plants in REGION_PLANTS.items():
        sub = selected[selected[MAT_PLANT].isin(plants)].copy()
        sub = sub.sort_values(
            [MAT_SLED, MAT_MATERIAL, MAT_BATCH], kind="mergesort"
        ).reset_index(drop=True)

        out = pd.DataFrame({col: sub.get(col) for col in OUTPUT_COLUMNS})
        out[MAT_SPECIAL_STOCK] = out[MAT_SPECIAL_STOCK].where(
            out[MAT_SPECIAL_STOCK].notna(), None
        )
        sheets[sheet_name] = out

    return sheets


# ---------------------------------------------------------------------------
# Excel export (matches the finished workbook: mm-dd-yy dates, plain stock)
# ---------------------------------------------------------------------------
_HEADER_FILL = PatternFill("solid", fgColor="FFF7F7F7")
_HEADER_FONT = Font(name="Arial", size=11, bold=True)
_BODY_FONT = Font(name="Arial", size=11)
_DATE_FMT = "mm-dd-yy"

_COL_WIDTHS = {
    MAT_MATERIAL: 13.0,
    MAT_DESCRIPTION: 36.0,
    MAT_PLANT: 6.14,
    MAT_PLANT_NAME: 16.0,
    OUT_BDM: 23.14,
    MAT_STORAGE_LOC: 18.57,
    MAT_STORAGE_DESC: 34.0,
    MAT_BATCH: 11.57,
    MAT_SLED: 27.43,
    OUT_LAST_SELL_DAY: 13.57,
    OUT_LAST_SELL_DT: 14.57,
    MAT_SPECIAL_STOCK: 33.29,
    MAT_UNRESTRICTED: 20.29,
    MAT_QUALITY: 28.43,
    MAT_BLOCKED: 15.71,
}

_TEXT_OUT_COLS = set(MAT_TEXT_COLS)


def _write_sheet(ws, df: pd.DataFrame) -> None:
    headers = list(df.columns)
    for c_idx, col in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=col)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(vertical="center")

    date_idx = {headers.index(c) + 1 for c in (MAT_SLED, OUT_LAST_SELL_DT) if c in headers}
    text_idx = {headers.index(c) + 1 for c in _TEXT_OUT_COLS if c in headers}

    for r_off, (_, row) in enumerate(df.iterrows()):
        r_idx = r_off + 2
        for c_idx, col in enumerate(headers, 1):
            val = row[col]
            if pd.isna(val):
                val = None
            if c_idx in text_idx and val is not None:
                val = str(val)
            if isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = _BODY_FONT
            if c_idx in date_idx and val is not None:
                cell.number_format = _DATE_FMT
            elif c_idx in text_idx:
                cell.number_format = "@"

    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{max(len(df) + 1, 1)}"
    for c_idx, col in enumerate(headers, 1):
        width = _COL_WIDTHS.get(col)
        if width:
            ws.column_dimensions[get_column_letter(c_idx)].width = width


def generate_excel(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Render the region DataFrames into a formatted workbook (bytes)."""
    wb = Workbook()
    wb.remove(wb.active)
    for sheet_name in REGION_PLANTS:
        ws = wb.create_sheet(sheet_name)
        _write_sheet(ws, sheets[sheet_name])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
