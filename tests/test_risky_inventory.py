"""Tests for the Risky Inventory engine.

Self-contained: every fixture workbook is built in memory, so these tests do not
depend on the supplied sample files.
"""
from __future__ import annotations

import io
from datetime import datetime

import openpyxl
import pytest

from src.risky_inventory_engine import (
    DETAIL_HEADERS,
    GRAND_TOTAL_LABEL,
    SUMMARY_HEADER,
    RiskyInventoryError,
    build_summary,
    generate_excel,
    load_detail,
    remove_duplicate_rows,
)


def _make_row(**over):
    """A full 20-column detail row keyed by header name; override any field."""
    base = {
        "Material": "10001334",
        "Material Description": "LC WHITE HOMINY",
        "Material Group": "D28",
        "Material Group Desc.": "LA COSTENA",
        "Purchasing Group": "104",
        "Description p. group": "Bita Farahani",
        "Brand Manager": "10",
        "Brand Manager Desc": "NANCY QUISPE PT",
        "MRP Area": "2920",
        "Batch": "OPAUG2726",
        "SLED Offset in days": -60,
        "Batch Expiry Date": datetime(2026, 8, 27),
        "MRP Last Sell Date": datetime(2026, 6, 28),
        "Quantity": 10,
        "Base Unit of Measure": "CS",
        "Qty Val. UoM": 10,
        "Total Stock": 10,
        "Moving price": 32.59,
        "Value": 325.90,
        "Batch Comment": None,
    }
    base.update(over)
    return [base[h] for h in DETAIL_HEADERS]


# Per-column number formats matching the real SAP export, keyed by header.
_FIXTURE_FORMATS = {
    "SLED Offset in days": "#,##0",
    "Batch Expiry Date": "mm-dd-yy",
    "MRP Last Sell Date": "mm-dd-yy",
    "Quantity": "#,##0.000",
    "Qty Val. UoM": "#,##0.000",
    "Total Stock": "#,##0.000",
    "Moving price": "#,##0.00",
    "Value": "#,##0.00",
}


def _make_xlsx(rows, *, sheet="Sheet1", headers=None):
    headers = headers if headers is not None else DETAIL_HEADERS
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(headers)
    for r in rows:
        ws.append(r)
    # Mirror the real export's column number formats so format passthrough is
    # exercised faithfully.
    for c, name in enumerate(headers, 1):
        fmt = _FIXTURE_FORMATS.get(name)
        if fmt:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, c).number_format = fmt
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# load_detail
# ---------------------------------------------------------------------------
def test_load_detail_reads_headers_and_rows():
    d = load_detail(_make_xlsx([_make_row(), _make_row(Batch="X2")]))
    assert d.headers == DETAIL_HEADERS
    assert len(d) == 2


def test_load_detail_rejects_missing_sheet1():
    wb = openpyxl.Workbook()
    wb.active.title = "Other"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    with pytest.raises(RiskyInventoryError):
        load_detail(buf)


def test_load_detail_rejects_wrong_headers():
    bad = ["Material", "Oops"] + DETAIL_HEADERS[2:]
    with pytest.raises(RiskyInventoryError):
        load_detail(_make_xlsx([_make_row()], headers=bad))


def test_load_detail_skips_blank_rows():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(DETAIL_HEADERS)
    ws.append(_make_row())
    ws.append([None] * len(DETAIL_HEADERS))  # trailing blank
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    assert len(load_detail(buf)) == 1


# ---------------------------------------------------------------------------
# remove_duplicate_rows
# ---------------------------------------------------------------------------
def test_already_cleaned_file_unchanged():
    d90 = load_detail(_make_xlsx([_make_row(Material="A"), _make_row(Material="B")]))
    d180 = load_detail(_make_xlsx([_make_row(Material="C"), _make_row(Material="D")]))
    clean = remove_duplicate_rows(d90, d180)
    assert clean.rows == d180.rows


def test_cumulative_file_drops_exactly_the_90day_rows_in_order():
    r90 = [_make_row(Material="A"), _make_row(Material="B")]
    r180_only = [_make_row(Material="C"), _make_row(Material="D"), _make_row(Material="E")]
    d90 = load_detail(_make_xlsx(r90))
    # cumulative repeats the 90-day rows first, then its own
    cumulative = load_detail(_make_xlsx(r90 + r180_only))
    clean = remove_duplicate_rows(d90, cumulative)
    expected = load_detail(_make_xlsx(r180_only))
    assert clean.rows == expected.rows  # order preserved, only own rows remain


def test_same_material_different_batch_not_treated_as_duplicate():
    r90 = [_make_row(Material="A", Batch="B1")]
    d90 = load_detail(_make_xlsx(r90))
    d180 = load_detail(_make_xlsx([
        _make_row(Material="A", Batch="B1"),   # exact dup -> removed
        _make_row(Material="A", Batch="B2"),   # same material, diff batch -> kept
    ]))
    clean = remove_duplicate_rows(d90, d180)
    assert len(clean) == 1
    assert clean.rows[0][DETAIL_HEADERS.index("Batch")] == "B2"


def test_int_float_whitespace_and_date_normalisation_match():
    d90 = load_detail(_make_xlsx([_make_row(Quantity=10, Material="A")]))
    # same row but quantity as float and a trailing space on a text field
    d180 = load_detail(_make_xlsx([
        _make_row(Quantity=10.0, Material="A ", )
    ]))
    clean = remove_duplicate_rows(d90, d180)
    assert len(clean) == 0


# ---------------------------------------------------------------------------
# build_summary
# ---------------------------------------------------------------------------
def test_summary_structure_grouping_sort_and_grand_total():
    rows = [
        _make_row(**{"Material Group Desc.": "ALPHA", "Quantity": 5, "Total Stock": 5, "Value": 100}),
        _make_row(**{"Material Group Desc.": "ALPHA", "Quantity": 5, "Total Stock": 5, "Value": 50}),
        _make_row(**{"Material Group Desc.": "BETA", "Quantity": 1, "Total Stock": 2, "Value": 300}),
    ]
    grid = build_summary(load_detail(_make_xlsx(rows)))

    # Three filter rows, blank, header row
    assert grid[0][0] == "Description p. group" and grid[0][1] == "(All)"
    assert grid[3] == [None] * len(SUMMARY_HEADER)
    assert grid[4] == SUMMARY_HEADER

    body = grid[5:-1]
    labels = [r[0] for r in body]
    assert labels == ["BETA", "ALPHA"]              # sorted by Sum of Value desc
    beta = body[0]
    assert (beta[5], beta[6], beta[7]) == (1, 2, 300)
    alpha = body[1]
    assert (alpha[5], alpha[6], alpha[7]) == (10, 10, 150)  # summed

    total = grid[-1]
    assert total[0] == GRAND_TOTAL_LABEL
    assert (total[5], total[6], total[7]) == (11, 12, 450)


def test_summary_value_tie_breaks_by_group_name_ascending():
    rows = [
        _make_row(**{"Material Group Desc.": "WESTKEY", "Value": 0, "Quantity": 1, "Total Stock": 1}),
        _make_row(**{"Material Group Desc.": "CARRS", "Value": 0, "Quantity": 1, "Total Stock": 1}),
    ]
    grid = build_summary(load_detail(_make_xlsx(rows)))
    labels = [r[0] for r in grid[5:-1]]
    assert labels == ["CARRS", "WESTKEY"]


# ---------------------------------------------------------------------------
# generate_excel
# ---------------------------------------------------------------------------
def test_generate_excel_sheet_order_and_detail_preserved():
    d90 = load_detail(_make_xlsx([_make_row(Material="A")]))
    d180 = load_detail(_make_xlsx([_make_row(Material="A"), _make_row(Material="Z")]))
    clean = remove_duplicate_rows(d90, d180)

    data = generate_excel(d90, clean)
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    assert wb.sheetnames == ["90D Detail", "90D Summary", "180D Detail", "180D Summary"]

    det = wb["180D Detail"]
    assert [det.cell(1, c).value for c in range(1, len(DETAIL_HEADERS) + 1)] == DETAIL_HEADERS
    # only the non-duplicate row remains
    assert det.max_row == 2
    assert det.cell(2, 1).value == "Z"


def test_generate_excel_applies_number_formats():
    d = load_detail(_make_xlsx([_make_row()]))
    data = generate_excel(d, d)
    wb = openpyxl.load_workbook(io.BytesIO(data))
    det = wb["90D Detail"]
    # Batch Expiry Date / Quantity / Value column formats
    assert det.cell(2, DETAIL_HEADERS.index("Batch Expiry Date") + 1).number_format == "mm-dd-yy"
    assert det.cell(2, DETAIL_HEADERS.index("Quantity") + 1).number_format == "#,##0.000"
    assert det.cell(2, DETAIL_HEADERS.index("Value") + 1).number_format == "#,##0.00"
    summ = wb["90D Summary"]
    assert summ.cell(6, 7).number_format == "#,##0"      # Sum of Total Stock
    assert summ.cell(6, 8).number_format == '"$"#,##0'   # Sum of Value
