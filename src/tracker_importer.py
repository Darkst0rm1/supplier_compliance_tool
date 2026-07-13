"""Read the approved-exception supplier list out of the Master Inbound Delivery
Compliance Tracker workbook.

The exceptions list is the UNION of two lists that Ops maintains by hand:
  1. Tracker sheet, Compliance Status == "NO -  Unable to Comply" (the Summary
     sheet calls these "Approved exceptions").
  2. "POs received" sheet, rows hand-marked "EXEMPT".

They overlap; "Unable to Comply" wins the reason because it is the
dropdown-driven, counted list.
"""
from __future__ import annotations

import pandas as pd

from .config import (
    REASON_EXEMPT_MARK,
    REASON_UNABLE_TO_COMPLY,
    TRACKER_EXEMPT_MARKER,
    TRACKER_EXEMPT_SHEET,
    TRACKER_NAME_COLUMN,
    TRACKER_SHEET,
    TRACKER_STATUS_COLUMN,
    TRACKER_STATUS_UNABLE_TO_COMPLY,
)
from .normalizer import normalize_supplier_name


class TrackerImportError(Exception):
    """The tracker workbook could not be read."""


def _find_column(df: pd.DataFrame, wanted: str) -> str | None:
    """Match a column tolerantly -- the workbook's headers carry stray spaces."""
    target = wanted.strip().casefold()
    for col in df.columns:
        if str(col).strip().casefold() == target:
            return col
    return None


def _seek_start(path_or_buffer) -> None:
    """Rewind a file-like object (e.g. a Streamlit UploadedFile) before reading.

    A plain path has no `seek`, so this is a no-op for it. A buffer is read
    twice by this module (once per sheet); without rewinding, the second read
    would see EOF and silently return nothing.
    """
    seek = getattr(path_or_buffer, "seek", None)
    if callable(seek):
        seek(0)


def _unable_to_comply(path_or_buffer) -> list[str]:
    _seek_start(path_or_buffer)
    try:
        df = pd.read_excel(path_or_buffer, sheet_name=TRACKER_SHEET)
    except ValueError as exc:
        raise TrackerImportError(
            f"The workbook has no '{TRACKER_SHEET}' sheet. Is this the Master "
            "Inbound Delivery Compliance Tracker?"
        ) from exc

    name_col = _find_column(df, TRACKER_NAME_COLUMN)
    status_col = _find_column(df, TRACKER_STATUS_COLUMN)
    if name_col is None or status_col is None:
        raise TrackerImportError(
            f"The '{TRACKER_SHEET}' sheet needs both a "
            f"'{TRACKER_NAME_COLUMN.strip()}' and a '{TRACKER_STATUS_COLUMN}' column."
        )

    status = df[status_col].fillna("").astype(str).str.strip()
    hit = status == TRACKER_STATUS_UNABLE_TO_COMPLY.strip()
    return [str(n).strip() for n in df.loc[hit, name_col].dropna()]


def _exempt_marked(path_or_buffer) -> list[str]:
    """Names on any row of 'POs received' carrying an EXEMPT marker.

    The marker sits in an unnamed column with no stable header, so scan every
    column of the row rather than relying on a position.
    """
    _seek_start(path_or_buffer)
    try:
        df = pd.read_excel(path_or_buffer, sheet_name=TRACKER_EXEMPT_SHEET)
    except ValueError:
        return []  # This sheet is optional; the Tracker sheet is not.

    if df.empty or len(df.columns) < 2:
        return []

    name_col = df.columns[0]
    others = df.columns[1:]
    marked = pd.Series(False, index=df.index)
    for col in others:
        cells = df[col].fillna("").astype(str).str.strip().str.upper()
        marked |= cells == TRACKER_EXEMPT_MARKER

    return [str(n).strip() for n in df.loc[marked, name_col].dropna() if str(n).strip()]


def read_tracker_exceptions(path_or_buffer) -> list[tuple[str, str]]:
    """Return de-duplicated (supplier_name, reason) pairs from the tracker.

    De-duplication is by normalized name; the first spelling seen wins, and
    "Unable to Comply" wins the reason because it is scanned first.
    """
    out: list[tuple[str, str]] = []
    seen: set[str] = set()

    for names, reason in (
        (_unable_to_comply(path_or_buffer), REASON_UNABLE_TO_COMPLY),
        (_exempt_marked(path_or_buffer), REASON_EXEMPT_MARK),
    ):
        for name in names:
            key = normalize_supplier_name(name)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append((name, reason))
    return out
