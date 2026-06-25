"""Processing engine for the Risky Inventory page.

Automates the one manual step in the existing Risky Inventory process. The user
downloads two reports:

1. a 90-day report, and
2. a cumulative 180-day report (which repeats the same first 90 days).

Normally the user opens the 180-day file and deletes by hand the rows that
already appear in the 90-day file. This module does only that: it removes from
the 180-day detail any row that is already present in the 90-day detail. Nothing
is renamed, recalculated, filtered, or added beyond what the supplied files
already contain.

The two supplied workbooks are the golden specification:

* ``Risky Inventory June 24 P2 - 90D.xlsx``
* ``Risky Inventory June 24 P2 - 180D.xlsx``

Each workbook has two sheets:

* ``Sheet1`` — the detailed inventory rows (20 columns, exact SAP export
  headers). Values are preserved verbatim: text codes stay text, dates stay
  dates, negatives / zeros / blanks are untouched.
* ``Sheet2`` — an Excel PivotTable summarising ``Sheet1``: one row per
  ``Material Group Desc.`` with Sum of Quantity / Sum of Total Stock / Sum of
  Value, sorted by Sum of Value descending, ending in a ``Grand Total`` row.
  Three report filters (``Description p. group``, ``Brand Manager Desc``,
  ``MRP Area``) sit above the table, all set to ``(All)``.

The summary is rebuilt from the (processed) detail rather than copied from the
pivot cache, so the cleaned 180-day summary reflects only the rows that remain.

This module is intentionally self-contained — like every other engine in this
app it stands alone and does not import the other report logic.
"""
from __future__ import annotations

import copy
import io
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


class RiskyInventoryError(Exception):
    """Raised when an uploaded file isn't a usable Risky Inventory export."""


# ---------------------------------------------------------------------------
# Source layout (exact sheet names and headers)
# ---------------------------------------------------------------------------
DETAIL_SHEET = "Sheet1"
SUMMARY_SHEET = "Sheet2"

# The full 20-column detail header, in order. The uploaded files must match.
DETAIL_HEADERS = [
    "Material",
    "Material Description",
    "Material Group",
    "Material Group Desc.",
    "Purchasing Group",
    "Description p. group",
    "Brand Manager",
    "Brand Manager Desc",
    "MRP Area",
    "Batch",
    "SLED Offset in days",
    "Batch Expiry Date",
    "MRP Last Sell Date",
    "Quantity",
    "Base Unit of Measure",
    "Qty Val. UoM",
    "Total Stock",
    "Moving price",
    "Value",
    "Batch Comment",
]

# Fields that together identify one inventory line. Comparing on these (rather
# than Material alone) keeps separate batches / inventory lines distinct, so an
# already-cleaned 180-day file loses nothing while a cumulative one loses
# exactly its first-90-day rows.
KEY_COLUMNS = [
    "Material",
    "Material Description",
    "Material Group Desc.",
    "Description p. group",
    "Brand Manager Desc",
    "MRP Area",
    "Batch",
    "SLED Offset in days",
    "Batch Expiry Date",
    "MRP Last Sell Date",
    "Quantity",
    "Qty Val. UoM",
    "Total Stock",
    "Moving price",
    "Value",
    "Batch Comment",
]

# ---------------------------------------------------------------------------
# Summary (Sheet2) layout — reproduced exactly from the supplied pivot tables
# ---------------------------------------------------------------------------
SUMMARY_FILTERS = [
    ("Description p. group", "(All)"),
    ("Brand Manager Desc", "(All)"),
    ("MRP Area", "(All)"),
]
SUMMARY_GROUP_COL = "Material Group Desc."
SUMMARY_HEADER = [
    "Material Group Desc.",
    "Material",
    "Material Description",
    "Batch",
    "Batch Expiry Date",
    "Sum of Quantity",
    "Sum of Total Stock",
    "Sum of Value",
]
GRAND_TOTAL_LABEL = "Grand Total"
# (detail column -> 0-based position in SUMMARY_HEADER) for the three value cols.
_SUMMARY_VALUE_COLS = [
    ("Quantity", 5),
    ("Total Stock", 6),
    ("Value", 7),
]
# Number formats and column widths copied from the supplied Sheet2 pivots.
SUMMARY_COL_WIDTHS = {
    "A": 52.71, "B": 15.71, "C": 18.43, "D": 13.43,
    "E": 19.71, "F": 15.71, "G": 18.43, "H": 13.43,
}
SUMMARY_NUMBER_FORMATS = {  # 1-based column index -> format
    6: "General",       # Sum of Quantity
    7: "#,##0",         # Sum of Total Stock
    8: '"$"#,##0',      # Sum of Value
}


# ---------------------------------------------------------------------------
# Detail table — values plus the formatting captured from the source sheet
# ---------------------------------------------------------------------------
@dataclass
class DetailTable:
    """The Sheet1 detail rows together with the formatting we read off it."""

    headers: list[str]
    rows: list[list[Any]]
    number_formats: dict[int, str] = field(default_factory=dict)   # 1-based col -> fmt
    column_widths: dict[str, float] = field(default_factory=dict)  # letter -> width
    header_fonts: list[Any] = field(default_factory=list)          # per-column copy
    header_fills: list[Any] = field(default_factory=list)
    header_alignments: list[Any] = field(default_factory=list)

    def __len__(self) -> int:
        return len(self.rows)

    @property
    def index(self) -> dict[str, int]:
        return {h: i for i, h in enumerate(self.headers)}


def load_detail(file_obj: Any) -> DetailTable:
    """Read ``Sheet1`` from an uploaded workbook, preserving native cell values.

    Text codes stay strings, dates stay ``datetime``, blanks stay blank. The
    per-column number formats, column widths and header cell styling are
    captured so the output workbook reproduces the source look exactly.
    """
    try:
        wb = load_workbook(file_obj, data_only=True)
    except Exception as exc:  # noqa: BLE001
        raise RiskyInventoryError(f"Could not open the workbook: {exc}") from exc

    if DETAIL_SHEET not in wb.sheetnames:
        raise RiskyInventoryError(
            f"Worksheet '{DETAIL_SHEET}' not found — this does not look like a "
            "Risky Inventory export."
        )
    ws = wb[DETAIL_SHEET]

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    # Trim trailing all-empty header columns.
    while headers and headers[-1] in (None, ""):
        headers.pop()
    if [h for h in headers] != DETAIL_HEADERS:
        raise RiskyInventoryError(
            "Sheet1 column headers do not match the expected Risky Inventory "
            "layout. Expected:\n  " + " | ".join(DETAIL_HEADERS)
        )

    ncols = len(headers)
    rows: list[list[Any]] = []
    for r in range(2, ws.max_row + 1):
        values = [ws.cell(r, c).value for c in range(1, ncols + 1)]
        if all(v is None for v in values):
            continue  # skip fully blank rows
        rows.append(values)

    # Capture per-column number formats from the first data row (column-uniform
    # in these exports). Fall back to the header cell if there are no rows.
    number_formats: dict[int, str] = {}
    fmt_row = 2 if ws.max_row >= 2 else 1
    for c in range(1, ncols + 1):
        number_formats[c] = ws.cell(fmt_row, c).number_format

    column_widths = {
        letter: dim.width
        for letter, dim in ws.column_dimensions.items()
        if dim.width is not None
    }

    header_fonts, header_fills, header_aligns = [], [], []
    for c in range(1, ncols + 1):
        cell = ws.cell(1, c)
        header_fonts.append(copy.copy(cell.font))
        header_fills.append(copy.copy(cell.fill))
        header_aligns.append(copy.copy(cell.alignment))

    return DetailTable(
        headers=headers,
        rows=rows,
        number_formats=number_formats,
        column_widths=column_widths,
        header_fonts=header_fonts,
        header_fills=header_fills,
        header_alignments=header_aligns,
    )


# ---------------------------------------------------------------------------
# De-duplication
# ---------------------------------------------------------------------------
def _norm(value: Any) -> str:
    """Normalise one cell for reliable row comparison (display is untouched)."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, bool):  # guard: bool is a subclass of int
        return str(value)
    if isinstance(value, (int, float)):
        return format(round(float(value), 6), ".6f")
    return str(value).strip()


def _row_key(row: list[Any], idx: dict[str, int]) -> tuple[str, ...]:
    return tuple(_norm(row[idx[col]]) for col in KEY_COLUMNS)


def remove_duplicate_rows(d90: DetailTable, d180: DetailTable) -> DetailTable:
    """Return a copy of the 180-day detail with rows already in the 90-day
    detail removed. Order and formatting of the 180-day file are preserved.

    If the 180-day file was already cleaned, nothing matches and every row is
    kept unchanged.
    """
    idx90 = d90.index
    key90 = {_row_key(r, idx90) for r in d90.rows}

    idx180 = d180.index
    kept = [r for r in d180.rows if _row_key(r, idx180) not in key90]

    cleaned = copy.copy(d180)
    cleaned.rows = kept
    return cleaned


# ---------------------------------------------------------------------------
# Summary (rebuilt from the processed detail, not the pivot cache)
# ---------------------------------------------------------------------------
def _num(value: Any) -> float:
    if isinstance(value, bool) or value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    return 0.0


def build_summary(detail: DetailTable) -> list[list[Any]]:
    """Build the Sheet2-style summary grid from a detail table.

    Returns the full sheet grid (filter rows, blank row, header, one row per
    ``Material Group Desc.`` sorted by Sum of Value descending, then a Grand
    Total row) — each row already the width of ``SUMMARY_HEADER``.
    """
    idx = detail.index
    g_i = idx[SUMMARY_GROUP_COL]
    val_src = {detail_col: idx[detail_col] for detail_col, _ in _SUMMARY_VALUE_COLS}

    totals: dict[str, list[float]] = {}
    for row in detail.rows:
        group = row[g_i]
        bucket = totals.setdefault(group, [0.0, 0.0, 0.0])
        bucket[0] += _num(row[val_src["Quantity"]])
        bucket[1] += _num(row[val_src["Total Stock"]])
        bucket[2] += _num(row[val_src["Value"]])

    # Sort by Sum of Value descending; ties broken by group name ascending
    # (matches the supplied pivots, e.g. CARRS before WESTKEY at value 0).
    ordered = sorted(totals.items(), key=lambda kv: (-kv[1][2], str(kv[0])))

    width = len(SUMMARY_HEADER)

    def _blank_row() -> list[Any]:
        return [None] * width

    grid: list[list[Any]] = []
    for label, value in SUMMARY_FILTERS:
        r = _blank_row()
        r[0], r[1] = label, value
        grid.append(r)
    grid.append(_blank_row())          # blank separator row
    grid.append(list(SUMMARY_HEADER))  # column header row

    grand = [0.0, 0.0, 0.0]
    for group, (q, ts, v) in ordered:
        r = _blank_row()
        r[0] = group
        r[5], r[6], r[7] = q, ts, v
        grid.append(r)
        grand[0] += q
        grand[1] += ts
        grand[2] += v

    total_row = _blank_row()
    total_row[0] = GRAND_TOTAL_LABEL
    total_row[5], total_row[6], total_row[7] = grand
    grid.append(total_row)

    return grid


# ---------------------------------------------------------------------------
# Workbook output
# ---------------------------------------------------------------------------
def _to_excel_value(value: Any) -> Any:
    """openpyxl accepts datetime, str, int, float and None directly."""
    return value


def _write_detail_sheet(ws, detail: DetailTable) -> None:
    headers = detail.headers
    for c, name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=name)
        if c - 1 < len(detail.header_fonts):
            cell.font = detail.header_fonts[c - 1]
            cell.fill = detail.header_fills[c - 1]
            cell.alignment = detail.header_alignments[c - 1]

    for r_off, row in enumerate(detail.rows):
        r = r_off + 2
        for c, value in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=_to_excel_value(value))
            fmt = detail.number_formats.get(c)
            if fmt:
                cell.number_format = fmt

    for letter, width in detail.column_widths.items():
        ws.column_dimensions[letter].width = width


def _write_summary_sheet(ws, grid: list[list[Any]]) -> None:
    for r_off, row in enumerate(grid):
        r = r_off + 1
        for c, value in enumerate(row, 1):
            if value is None:
                continue
            cell = ws.cell(row=r, column=c, value=value)
            fmt = SUMMARY_NUMBER_FORMATS.get(c)
            if fmt:
                cell.number_format = fmt
    for letter, width in SUMMARY_COL_WIDTHS.items():
        ws.column_dimensions[letter].width = width


def generate_excel(d90: DetailTable, d180_clean: DetailTable) -> bytes:
    """Render the four-sheet Risky Inventory workbook (bytes).

    Sheets, in order: ``90D Detail``, ``90D Summary``, ``180D Detail``,
    ``180D Summary``. The 180-day sheets use the cleaned detail.
    """
    wb = Workbook()
    wb.remove(wb.active)

    _write_detail_sheet(wb.create_sheet("90D Detail"), d90)
    _write_summary_sheet(wb.create_sheet("90D Summary"), build_summary(d90))
    _write_detail_sheet(wb.create_sheet("180D Detail"), d180_clean)
    _write_summary_sheet(wb.create_sheet("180D Summary"), build_summary(d180_clean))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
