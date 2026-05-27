"""Load, validate, and month-filter the Portal export.

Tolerates the richer portal export shape seen on the real Inbound Delivery
File List page (columns like File Name, File Status, Invalid Comment) while
still accepting the minimal three-column template for back-compat.
"""
from __future__ import annotations

import pandas as pd

from .config import (
    PORTAL_COLUMN_ALIASES,
    PORTAL_OPTIONAL_COLUMNS,
    PORTAL_REQUIRED_COLUMNS,
)
from .normalizer import split_multi_po


class PortalImportError(Exception):
    """Raised when the Portal file is missing columns or has unparseable dates."""


def load_portal(file, report_year: int, report_month: int) -> pd.DataFrame:
    """Read the Portal export and return rows filtered to the selected month.

    Each portal row is exploded so that one row in the result corresponds to one
    normalized PO. The original cell value is preserved in 'PO Number' for audit.
    """
    try:
        df = pd.read_excel(file, dtype=str)
    except Exception as e:  # pragma: no cover - bubbled to UI
        raise PortalImportError(f"Could not read Portal file: {e}") from e

    # Rename real portal headers (PO Number(s), Supplier) to canonical names.
    df = df.rename(columns=PORTAL_COLUMN_ALIASES).copy()

    missing = [c for c in PORTAL_REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise PortalImportError(
            "Portal file is missing required columns: " + ", ".join(missing)
        )

    # Add any optional columns that didn't appear in this export so the engine
    # always sees the same shape.
    for col in PORTAL_OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df["Upload Date"] = pd.to_datetime(df["Upload Date"], errors="coerce")
    if df["Upload Date"].isna().all():
        raise PortalImportError(
            "Upload Date column could not be parsed for any row."
        )

    # Filter to the selected report month.
    mask = (
        (df["Upload Date"].dt.year == report_year)
        & (df["Upload Date"].dt.month == report_month)
    )
    df = df.loc[mask].copy()

    # Explode multi-PO cells (the real portal exports comma-separated POs).
    df["__pos"] = df["PO Number"].apply(split_multi_po)
    df = df.explode("__pos", ignore_index=True)
    df["Normalized PO Number"] = df["__pos"].fillna("")
    df = df.drop(columns="__pos")

    # Clean string columns (don't touch dates).
    for col in ["Supplier Name", "File Name", "Uploaded By", "File Status",
                "Downloaded By", "Invalid Comment"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    # Parse Download Date if present.
    if "Download Date" in df.columns:
        df["Download Date"] = pd.to_datetime(df["Download Date"], errors="coerce")

    return df
