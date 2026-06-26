"""Load and normalize the SAP receiving-history export for batch-quality review.

The export is one row per receiving line on the sheet ``SAPUI5 Export`` (~30
columns, exact SAP headers). We read everything as text, normalize the headers
to canonical names internally while preserving the original business-friendly
labels for display, keep id columns as clean text strings (trailing ``.0``
removed, legitimate leading zeros preserved), parse the two date columns, and
add a comparison-only ``Normalized Batch`` field.

The real export uses ``Purchase order`` / ``Material Desc`` / ``Total SLED`` /
``Remain SLED`` / ``Min SLED`` (not the idealized names in the spec). The alias
map below tolerates either spelling so a re-exported file with slightly
different headers still loads.
"""
from __future__ import annotations

import re
from typing import Any

import pandas as pd


class BatchQualityError(Exception):
    """Raised when an uploaded file isn't a usable SAP receiving export."""


SAP_SHEET = "SAPUI5 Export"

# ---------------------------------------------------------------------------
# Canonical column names (the names used everywhere downstream + on display)
# ---------------------------------------------------------------------------
COL_VENDOR = "Vendor"
COL_VENDOR_NAME = "Vendor Name"
COL_SUPPLIER = "Supplier"
COL_SUPPLIER_NAME = "Supplier Name"
COL_PO = "Purchase order"
COL_PLANT = "Plant"
COL_PURCH_GROUP = "Purchasing Group"
COL_MATERIAL = "Material"
COL_MATERIAL_DESC = "Material Desc"
COL_BATCH = "Batch"
COL_BATCH_SLED = "Batch SLED"
COL_RECEIVED = "Received Date"
COL_QTY = "Qty Base"

NORMALIZED_BATCH = "Normalized Batch"

# Treated as text (trailing .0 stripped, leading zeros preserved).
TEXT_COLS = [
    COL_VENDOR, COL_VENDOR_NAME, COL_SUPPLIER, COL_SUPPLIER_NAME,
    COL_PO, COL_PLANT, COL_PURCH_GROUP, COL_MATERIAL, COL_MATERIAL_DESC, COL_BATCH,
]
# Batch keeps its surrounding whitespace verbatim (Rule 5 looks for it).
TEXT_COLS_KEEP_SPACES = {COL_BATCH}

DATE_COLS = [COL_BATCH_SLED, COL_RECEIVED]
NUMERIC_COLS = [COL_QTY]

REQUIRED_COLS = [COL_MATERIAL, COL_BATCH]


def _norm_header(h: Any) -> str:
    return re.sub(r"\s+", " ", str(h).strip()).lower()


# Map of normalized incoming header -> canonical name. Built from the canonical
# names themselves plus tolerated aliases (spec spellings, common variants).
_ALIASES: dict[str, str] = {}
for _c in [
    COL_VENDOR, COL_VENDOR_NAME, COL_SUPPLIER, COL_SUPPLIER_NAME, COL_PO,
    COL_PLANT, COL_PURCH_GROUP, COL_MATERIAL, COL_MATERIAL_DESC, COL_BATCH,
    COL_BATCH_SLED, COL_RECEIVED, COL_QTY,
]:
    _ALIASES[_norm_header(_c)] = _c
_ALIASES.update({
    "purchase order": COL_PO,
    "purchase order(s)": COL_PO,
    "po": COL_PO,
    "po number": COL_PO,
    "material description": COL_MATERIAL_DESC,
    "material desc.": COL_MATERIAL_DESC,
    "supplier name": COL_SUPPLIER_NAME,
    "vendor name": COL_VENDOR_NAME,
    "quantity": COL_QTY,
    "qty": COL_QTY,
})


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename recognized headers to canonical names; leave others untouched.

    Matching is case- and whitespace-insensitive so a re-export with slightly
    different header casing/spacing still maps. The first column to claim a
    canonical name wins (avoids collisions when both spellings are present)."""
    rename: dict[Any, str] = {}
    claimed: set[str] = set()
    for col in df.columns:
        canonical = _ALIASES.get(_norm_header(col))
        if canonical and canonical not in claimed and col != canonical:
            rename[col] = canonical
            claimed.add(canonical)
        elif canonical:
            claimed.add(canonical)
    return df.rename(columns=rename)


def normalize_batch(value: Any) -> str:
    """Comparison-only normalization: uppercase, drop spaces and separators
    (hyphens/slashes/periods/underscores and any other punctuation), keep
    letters and digits. ``31-5357`` / ``31 5357`` / ``31/5357`` -> ``315357``.
    Never replaces the original Batch value."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"[^0-9A-Z]", "", str(value).strip().upper())


def _clean_text(series: pd.Series, keep_spaces: bool) -> pd.Series:
    # fillna("") first: pandas 3.0 astype(str) leaves NaN as a float rather than
    # the string "nan", so a later string .replace would never clear missing.
    out = (
        series.fillna("").astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .replace({"nan": "", "None": "", "NaT": "", "<NA>": ""})
    )
    if not keep_spaces:
        out = out.str.strip()
    return out


def load_batch_data(file_obj: Any) -> pd.DataFrame:
    """Read the SAP receiving export and return a normalized DataFrame.

    Reads the ``SAPUI5 Export`` sheet (falls back to the first sheet) as text,
    canonicalizes headers, cleans text ids, parses dates, coerces quantities to
    numeric, and adds ``Normalized Batch``. Original values are otherwise
    preserved for the detailed output.
    """
    file_obj.seek(0)
    try:
        xl = pd.ExcelFile(file_obj, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001
        raise BatchQualityError(f"Could not open the workbook: {exc}") from exc

    sheet = SAP_SHEET if SAP_SHEET in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = normalize_columns(df)

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise BatchQualityError(
            "This doesn't look like a SAP receiving export — missing column(s): "
            + ", ".join(missing)
        )

    for c in TEXT_COLS:
        if c in df.columns:
            df[c] = _clean_text(df[c], keep_spaces=c in TEXT_COLS_KEEP_SPACES)

    for c in DATE_COLS:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    for c in NUMERIC_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    df[NORMALIZED_BATCH] = df[COL_BATCH].map(normalize_batch)
    return df.reset_index(drop=True)
