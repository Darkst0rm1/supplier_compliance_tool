"""Load, validate, and month-filter the dock Receiving Log.

The receiving log is an optional third input. It is a hand-maintained
operational workbook, not a system export, which drives every decision here:

 - One sheet per month-range, and the schema changed mid-year. The Jan-Apr
   sheets predate the Y/N audit columns, so they are skipped rather than read
   as all-blank.
 - The header row is not row 0 -- a title banner and a "Problem Key" note sit
   above it -- so the header row is located by content, not by position.
 - Headers drift in spacing ("Correct BBD Received Y/N" vs "... Y / N"), so
   aliases are matched on a whitespace-stripped lowercase form.
 - PO cells are hand-typed and may hold several POs separated by "/", ",",
   "&" or spaces, and may hold non-SAP references (TR-34306, GHPO-23467,
   "Return"). Those tokens simply never match a SAP PO and are dropped.
 - Rows dated in the future are pre-formatted appointment slots with nothing
   filled in; the month filter removes them naturally.
"""
from __future__ import annotations

import re

import pandas as pd

from .config import (
    EXCLUDED_PO_PREFIXES,
    RECEIVING_AUDIT_COLUMNS,
    RECEIVING_CANONICAL_COLUMNS,
    RECEIVING_COLUMN_ALIASES,
    RECEIVING_MIN_PO_DIGITS,
    RECEIVING_NO,
    RECEIVING_NO_VALUES,
    RECEIVING_YES,
    RECEIVING_YES_VALUES,
)
from .normalizer import is_excluded_po, split_multi_po

# How far down a sheet to look for the header row before giving up.
_HEADER_SCAN_ROWS = 15

_WS_RE = re.compile(r"\s+")


class ReceivingImportError(Exception):
    """Raised when no sheet in the Receiving Log carries the audit columns."""


def _header_key(value) -> str:
    """Whitespace-stripped lowercase form used to match a header to an alias."""
    if value is None:
        return ""
    s = str(value)
    if s.lower() == "nan":
        return ""
    return _WS_RE.sub("", s).strip().lower()


def normalize_yes_no(value) -> str:
    """Map a free-text Y/N cell to 'YES', 'NO', or '' when unanswered."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    s = _WS_RE.sub(" ", str(value)).strip().upper()
    if not s or s == "NAN":
        return ""
    if s in RECEIVING_YES_VALUES:
        return RECEIVING_YES
    if s in RECEIVING_NO_VALUES:
        return RECEIVING_NO
    return ""


def _looks_like_po(token: str) -> bool:
    """True if a token could be a SAP PO number (purely numeric, long enough)."""
    return bool(token) and token.isdigit() and len(token) >= RECEIVING_MIN_PO_DIGITS


def _find_header_row(raw: pd.DataFrame) -> int | None:
    """Return the index of the row that holds the column headers, or None.

    The header row is the one containing a PO column *and* a Date column --
    both are present in every version of the log's schema.
    """
    limit = min(_HEADER_SCAN_ROWS, len(raw))
    for i in range(limit):
        keys = {_header_key(v) for v in raw.iloc[i].tolist()}
        has_po = bool(keys & {"po#", "po", "ponumber", "ponumber(s)"})
        has_date = "date" in keys
        if has_po and has_date:
            return i
    return None


def _canonicalize(df: pd.DataFrame) -> pd.DataFrame:
    """Rename recognized headers to canonical names and drop the rest.

    Duplicate canonical names (the real log has two unnamed spacer columns and
    can repeat a header) keep the first occurrence only.
    """
    renamed: dict[int, str] = {}
    seen: set[str] = set()
    for pos, col in enumerate(df.columns):
        canon = RECEIVING_COLUMN_ALIASES.get(_header_key(col))
        if canon and canon not in seen:
            seen.add(canon)
            renamed[pos] = canon

    out = df.iloc[:, list(renamed.keys())].copy()
    out.columns = list(renamed.values())
    return out


def _read_sheet(file, sheet_name: str) -> pd.DataFrame | None:
    """Read one sheet into canonical shape, or None if it isn't a usable log."""
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None, dtype=str)
    if raw.empty:
        return None

    header_row = _find_header_row(raw)
    if header_row is None:
        return None

    df = raw.iloc[header_row + 1:].copy()
    df.columns = raw.iloc[header_row]
    df = _canonicalize(df)

    # Sheets on the pre-audit schema (Jan-Apr) are not usable for this report.
    if not any(c in df.columns for c in RECEIVING_AUDIT_COLUMNS):
        return None

    df["Source Sheet"] = sheet_name
    return df.dropna(how="all")


def load_receiving(file, report_year: int, report_month: int) -> pd.DataFrame:
    """Read the Receiving Log and return rows for the selected month.

    Each row is exploded so one output row corresponds to one normalized PO.
    The result carries an ``attrs`` dict describing what was skipped, so the
    UI can be honest about coverage instead of silently under-reporting.
    """
    try:
        book = pd.ExcelFile(file)
        sheet_names = book.sheet_names
    except Exception as e:  # pragma: no cover - bubbled to UI
        raise ReceivingImportError(f"Could not read Receiving Log: {e}") from e

    frames: list[pd.DataFrame] = []
    skipped: list[str] = []
    for name in sheet_names:
        try:
            part = _read_sheet(book, name)
        except Exception:  # a stray sheet must not sink the whole import
            part = None
        if part is None or part.empty:
            skipped.append(name)
        else:
            frames.append(part)

    if not frames:
        raise ReceivingImportError(
            "No sheet in this workbook has the receiving audit columns "
            "(Inbound File Y/N, Correct Batch/BBD/QTY Received). Sheets "
            "checked: " + ", ".join(sheet_names)
        )

    df = pd.concat(frames, ignore_index=True)

    # Guarantee a stable shape regardless of which columns a sheet carried.
    for col in RECEIVING_CANONICAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df["Receiving Date"] = pd.to_datetime(df["Receiving Date"], errors="coerce")
    total_rows = len(df)

    # Scope to the report month. Future-dated rows are unfilled appointment
    # slots and fall away here along with every other out-of-month row.
    mask = (
        (df["Receiving Date"].dt.year == report_year)
        & (df["Receiving Date"].dt.month == report_month)
    )
    df = df.loc[mask].copy()
    rows_in_month = len(df)

    # Normalize the four Y/N answers before anything downstream reads them.
    for col in RECEIVING_AUDIT_COLUMNS:
        df[col] = df[col].apply(normalize_yes_no)

    for col in ["Carrier", "Results of Inspection", "Receiver Initials", "Comments"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # Rows with no PO at all can never be joined -- count them so the UI can
    # report the real coverage of the log.
    rows_without_po = int(
        (df["PO Number"].fillna("").astype(str).str.strip() == "").sum()
    )

    df["__pos"] = df["PO Number"].apply(split_multi_po)
    df = df.explode("__pos", ignore_index=True)
    df["Normalized PO Number"] = df["__pos"].fillna("")
    df = df.drop(columns="__pos")

    # Drop tokens that cannot be SAP POs (carrier refs, free text, short
    # internal numbers). Counted so the UI can report real coverage.
    non_po_refs: list[str] = []
    if not df.empty:
        is_po = df["Normalized PO Number"].apply(_looks_like_po).astype(bool)
        non_po_refs = sorted(
            set(df.loc[~is_po & (df["Normalized PO Number"] != ""),
                       "Normalized PO Number"])
        )
        df = df[is_po].copy()

    # Apply the same excluded-PO-type rule as the SAP and portal sides.
    # Guard the empty case: .apply on an empty frame yields an object Series
    # whose .sum() is "" rather than 0.
    excluded_count = 0
    if not df.empty:
        excluded_mask = df["Normalized PO Number"].apply(
            lambda po: is_excluded_po(po, EXCLUDED_PO_PREFIXES)
        )
        excluded_count = int(excluded_mask.sum())
        if excluded_count:
            df = df.loc[~excluded_mask].copy()

    df.attrs["skipped_sheets"] = skipped
    df.attrs["total_rows"] = total_rows
    df.attrs["rows_in_month"] = rows_in_month
    df.attrs["rows_without_po"] = rows_without_po
    df.attrs["non_po_references"] = non_po_refs
    df.attrs["excluded_po_count"] = excluded_count
    return df


def build_po_lookup(receiving_df: pd.DataFrame) -> pd.DataFrame:
    """Collapse the log to one row per PO, keeping the first real answer.

    A PO can be received across several dock rows (split loads, re-checks).
    Taking the first *answered* value per column -- rather than the first row
    -- avoids letting a blank row erase an answer recorded elsewhere.
    """
    cols = RECEIVING_AUDIT_COLUMNS + [
        "Carrier", "Receiving Date", "Results of Inspection", "Comments",
    ]
    if receiving_df.empty:
        return pd.DataFrame(columns=cols)

    def _first_answer(series: pd.Series):
        for v in series:
            if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip():
                return v
        return "" if series.name in RECEIVING_AUDIT_COLUMNS else series.iloc[0]

    present = [c for c in cols if c in receiving_df.columns]
    return receiving_df.groupby("Normalized PO Number")[present].agg(_first_answer)
