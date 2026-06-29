"""Tests for the Risky Inventory engine.

Self-contained: every fixture workbook is built in memory, so these tests do not
depend on the supplied sample files.
"""
from __future__ import annotations

import io
import zipfile
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pytest

from src.risky_inventory_engine import (
    BUCKET_0_90,
    BUCKET_91_180,
    BUCKET_NONE,
    DETAIL_HEADERS,
    TEMPLATE_PATH,
    RiskyInventoryError,
    assign_buckets,
    bucket_for,
    compute_cutoff,
    generate_excel,
    load_detail,
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
# generate_excel — fills the PivotTable template
# ---------------------------------------------------------------------------
def test_generate_excel_fills_detail_and_keeps_pivot():
    rows = [
        _make_row(Material="A", **{"MRP Last Sell Date": datetime(2026, 8, 1)}),
        _make_row(Material="B", **{"MRP Last Sell Date": datetime(2026, 12, 1)}),
    ]
    detail = load_detail(_make_xlsx(rows))
    bucketed, _ = assign_buckets(detail, compute_cutoff(date(2026, 6, 24)))
    data = generate_excel(bucketed)

    # Pivot parts survive (real, refreshable PivotTable).
    parts = [n for n in zipfile.ZipFile(io.BytesIO(data)).namelist() if "pivot" in n.lower()]
    assert len(parts) == 5

    wb = openpyxl.load_workbook(io.BytesIO(data))
    det = wb["Detail"]
    assert det.cell(1, 21).value == "Bucket"
    assert det.max_row == 3                      # header + 2 rows
    assert det.cell(2, 21).value == "0-90 Day"
    assert det.cell(3, 21).value == "91-180 Day"
    piv = wb["Summary"]._pivots[0]
    assert piv.cache.refreshOnLoad is True
    assert piv.cache.cacheSource.worksheetSource.ref == "A1:U3"


def test_generate_excel_applies_number_formats():
    rows = [_make_row(Material="A", **{"MRP Last Sell Date": datetime(2026, 8, 1)})]
    detail = load_detail(_make_xlsx(rows))
    bucketed, _ = assign_buckets(detail, compute_cutoff(date(2026, 6, 24)))
    data = generate_excel(bucketed)
    wb = openpyxl.load_workbook(io.BytesIO(data))
    det = wb["Detail"]
    assert det.cell(2, DETAIL_HEADERS.index("Batch Expiry Date") + 1).number_format == "mm-dd-yy"
    assert det.cell(2, DETAIL_HEADERS.index("Quantity") + 1).number_format == "#,##0.000"
    assert det.cell(2, DETAIL_HEADERS.index("Value") + 1).number_format == "#,##0.00"


# ---------------------------------------------------------------------------
# Bucketing
# ---------------------------------------------------------------------------
def test_compute_cutoff_is_run_date_plus_90():
    assert compute_cutoff(date(2026, 6, 24)) == date(2026, 9, 22)


def test_bucket_for_boundaries():
    cutoff = date(2026, 9, 22)
    assert bucket_for(datetime(2026, 9, 22), cutoff) == BUCKET_0_90   # inclusive
    assert bucket_for(datetime(2026, 9, 21), cutoff) == BUCKET_0_90
    assert bucket_for(datetime(2026, 9, 23), cutoff) == BUCKET_91_180
    assert bucket_for(None, cutoff) == BUCKET_NONE
    assert bucket_for("", cutoff) == BUCKET_NONE


def test_assign_buckets_appends_column_and_counts():
    rows = [
        _make_row(Material="A", **{"MRP Last Sell Date": datetime(2026, 8, 1)}),   # 0-90
        _make_row(Material="B", **{"MRP Last Sell Date": datetime(2026, 12, 1)}),  # 91-180
        _make_row(Material="C", **{"MRP Last Sell Date": None}),                   # none
    ]
    detail = load_detail(_make_xlsx(rows))
    bucketed, counts = assign_buckets(detail, compute_cutoff(date(2026, 6, 24)))
    assert bucketed.headers[-1] == "Bucket"
    assert [r[-1] for r in bucketed.rows] == [BUCKET_0_90, BUCKET_91_180, BUCKET_NONE]
    assert counts == {BUCKET_0_90: 1, BUCKET_91_180: 1, BUCKET_NONE: 1}
    assert len(bucketed.rows) == 3 and len(bucketed.rows[0]) == len(detail.headers) + 1


# ---------------------------------------------------------------------------
# Template asset structure
# ---------------------------------------------------------------------------
def test_template_asset_is_valid_and_data_clean():
    assert Path(TEMPLATE_PATH).exists()
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    assert set(wb.sheetnames) == {"Detail", "Summary"}
    hdr = [wb["Detail"].cell(1, c).value for c in range(1, 22)]
    assert hdr[:20] == DETAIL_HEADERS and hdr[20] == "Bucket"
    piv = wb["Summary"]._pivots[0]
    names = [f.name for f in piv.cache.cacheFields]
    bucket_idx = names.index("Bucket")
    assert bucket_idx in [pf.fld for pf in piv.pageFields]       # Bucket is a page filter
    assert piv.cache.refreshOnLoad is True
    # No real supplier data committed in the cache definition.
    cdef = zipfile.ZipFile(TEMPLATE_PATH).read(
        "xl/pivotCache/pivotCacheDefinition1.xml").decode()
    assert "10001334" not in cdef
