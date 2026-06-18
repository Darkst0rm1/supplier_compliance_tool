"""Load and validate the SAP export.

Real SAP exports use headers like ``Plant``, ``Vendor``, ``Appt. Date``.
This module renames them to the canonical names used by the engine, and
tolerates older exports that omit Vendor Number / Warehouse entirely.
"""
from __future__ import annotations

import pandas as pd

from .config import (
    EXCLUDED_PO_PREFIXES,
    SAP_CANONICAL_COLUMNS,
    SAP_COLUMN_ALIASES,
    SAP_HARD_REQUIRED_COLUMNS,
    SAP_OPTIONAL_DATE_COLUMNS,
)
from .normalizer import is_excluded_po, normalize_po


class SapImportError(Exception):
    """Raised when the SAP file is unreadable or missing PO Number / Inbound Delivery."""


_STRING_COLS = [
    "Vendor Number",
    "Vendor Name",
    "Warehouse",
    "PO Status",
    "Inbound Delivery",
    "Inbound Delivery Status",
]
_DATE_COLS = ["Appointment Date", "Delivery Date"]


def load_sap(file) -> pd.DataFrame:
    """Read the SAP export, normalize headers, and return a canonical dataframe."""
    try:
        df = pd.read_excel(file, dtype=str)
    except Exception as e:  # pragma: no cover - bubbled to UI
        raise SapImportError(f"Could not read SAP file: {e}") from e

    # 1) Map real SAP headers -> canonical headers (Plant -> Warehouse, etc).
    df = df.rename(columns=SAP_COLUMN_ALIASES).copy()

    # 2) Hard requirement: we must at least have a PO Number and Inbound Delivery
    #    column, otherwise compliance logic has nothing to chew on.
    missing_required = [c for c in SAP_HARD_REQUIRED_COLUMNS if c not in df.columns]
    if missing_required:
        raise SapImportError(
            "SAP file is missing required columns: " + ", ".join(missing_required)
        )

    # 3) If the export lacks a separate PO Status column (most do — the C/A/P
    #    codes live in Inbound Delivery Status), source it from there.
    if "PO Status" not in df.columns:
        if "Inbound Delivery Status" in df.columns:
            df["PO Status"] = df["Inbound Delivery Status"]
        else:
            df["PO Status"] = ""

    # 4) Add any other canonical columns that didn't show up in this export so
    #    every downstream consumer sees a consistent shape.
    for col in SAP_CANONICAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # 5) Normalize PO Numbers and tidy string columns.
    df["Normalized PO Number"] = df["PO Number"].apply(normalize_po)

    # 5a) Disregard excluded PO types (e.g. POs starting with "6") so they
    #     never feed compliance, rollups, or bill-back. Record the count.
    excluded_mask = df["Normalized PO Number"].apply(
        lambda po: is_excluded_po(po, EXCLUDED_PO_PREFIXES)
    )
    excluded_count = int(excluded_mask.sum())
    if excluded_count:
        df = df.loc[~excluded_mask].copy()
    df.attrs["excluded_po_count"] = excluded_count

    for col in _STRING_COLS:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df["PO Status"] = df["PO Status"].str.upper()

    # 6) Parse dates leniently.
    for col in _DATE_COLS:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    # 7) Optional pickup-date columns: parse if present, else add blank so the
    #    SAP Export Data sheet has a consistent shape.
    for col in SAP_OPTIONAL_DATE_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        else:
            df[col] = pd.NaT

    return df


def describe_missing_optionals(df: pd.DataFrame) -> list[str]:
    """Return a list of canonical columns that loaded entirely blank.

    Useful for surfacing soft warnings in the UI (e.g. "Warehouse column was
    empty in this export — Warehouse Summary will be empty").
    """
    notes: list[str] = []
    for col in ("Vendor Number", "Warehouse"):
        if col in df.columns and (df[col].astype(str).str.strip() == "").all():
            notes.append(col)
    return notes
