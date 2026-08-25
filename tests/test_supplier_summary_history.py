"""Tests for the Supplier Summary history store."""
from __future__ import annotations

import os
from datetime import date, datetime

import pandas as pd
import pytest

from src.supplier_summary_history import summary_to_snapshot_rows


def _summary(rows):
    """A Supplier Summary-shaped DataFrame. ``rows`` are dicts with at least
    the required columns; anything omitted defaults sensibly."""
    defaults = {
        "Vendor Number": "0", "Vendor Name": "", "BDM": "", "BDM Description": "",
        "Exception Status": "Expected", "Total SAP POs": 0,
        "SAP POs With Inbound Delivery": 0, "Portal Files Found": 0,
        "Missing Portal Files": 0, "Invalid Portal Uploads": 0,
        "Portal Files With No SAP Inbound": 0, "Closed POs": 0, "Processing POs": 0,
    }
    return pd.DataFrame([{**defaults, **r} for r in rows])


# ---------------------------------------------------------------------------
# summary_to_snapshot_rows — pure transform
# ---------------------------------------------------------------------------
def test_basic_conversion():
    df = _summary([{
        "Vendor Number": "1001", "Vendor Name": "Acme Foods", "BDM": "J DOE",
        "Total SAP POs": 10, "SAP POs With Inbound Delivery": 8,
        "Portal Files Found": 6, "Missing Portal Files": 2,
        "Invalid Portal Uploads": 1,
    }])
    rows = summary_to_snapshot_rows(df, date(2026, 8, 26), 2026, 8)
    assert len(rows) == 1
    r = rows[0]
    assert r["snapshot_date"] == date(2026, 8, 26)
    assert r["report_year"] == 2026
    assert r["report_month"] == 8
    assert r["vendor_number"] == "1001"
    assert r["vendor_name"] == "Acme Foods"
    assert r["bdm"] == "J DOE"
    assert r["sap_pos_with_inbound"] == 8
    assert r["portal_files_found"] == 6
    assert r["missing_portal_files"] == 2
    assert r["invalid_portal_uploads"] == 1
    assert r["compliance_pct"] == pytest.approx(6 / 8)


def test_zero_inbound_gives_none_compliance_pct():
    df = _summary([{
        "Vendor Number": "1002", "Vendor Name": "No Inbound Co",
        "Total SAP POs": 3, "SAP POs With Inbound Delivery": 0,
        "Portal Files Found": 0,
    }])
    rows = summary_to_snapshot_rows(df, date(2026, 8, 26), 2026, 8)
    assert rows[0]["compliance_pct"] is None


def test_missing_required_column_raises():
    df = pd.DataFrame([{"Vendor Number": "1001"}])
    with pytest.raises(ValueError, match="missing column"):
        summary_to_snapshot_rows(df, date(2026, 8, 26), 2026, 8)


def test_snapshot_date_accepts_a_datetime():
    df = _summary([{"Vendor Number": "1001", "Vendor Name": "Acme"}])
    rows = summary_to_snapshot_rows(df, datetime(2026, 8, 26, 14, 30), 2026, 8)
    assert rows[0]["snapshot_date"] == date(2026, 8, 26)


def test_multiple_vendors_each_get_a_row():
    df = _summary([
        {"Vendor Number": "1001", "Vendor Name": "A", "SAP POs With Inbound Delivery": 5, "Portal Files Found": 5},
        {"Vendor Number": "1002", "Vendor Name": "B", "SAP POs With Inbound Delivery": 4, "Portal Files Found": 2},
    ])
    rows = summary_to_snapshot_rows(df, date(2026, 8, 26), 2026, 8)
    assert [r["vendor_number"] for r in rows] == ["1001", "1002"]
    assert rows[0]["compliance_pct"] == 1.0
    assert rows[1]["compliance_pct"] == 0.5


def test_blank_optional_text_columns_become_empty_string():
    df = _summary([{"Vendor Number": "1001", "Vendor Name": "A", "BDM": None}])
    rows = summary_to_snapshot_rows(df, date(2026, 8, 26), 2026, 8)
    assert rows[0]["bdm"] == ""


# ---------------------------------------------------------------------------
# SupplierSummaryHistoryStore — needs a real Postgres (skipped otherwise)
# ---------------------------------------------------------------------------
from src.supplier_summary_history import SupplierSummaryHistoryStore  # noqa: E402

TEST_DSN = os.environ.get("TEST_DATABASE_URL")
needs_db = pytest.mark.skipif(not TEST_DSN, reason="TEST_DATABASE_URL not set")
PREFIX = "PYTEST_SS_"


@pytest.fixture
def store():
    s = SupplierSummaryHistoryStore(TEST_DSN)
    s.ensure_schema()
    yield s
    with s._connect() as conn:  # noqa: SLF001 - test cleanup
        conn.execute(
            "DELETE FROM supplier_summary_snapshots WHERE vendor_number LIKE %s",
            (PREFIX + "%",),
        )


def _rows(snapshot_date, vendor_number, **overrides):
    df = _summary([{
        "Vendor Number": vendor_number, "Vendor Name": "Test Vendor",
        "SAP POs With Inbound Delivery": 10, "Portal Files Found": 7,
        **overrides,
    }])
    return summary_to_snapshot_rows(df, snapshot_date, 2026, 8)


@needs_db
def test_save_and_list_dates(store):
    store.save_snapshot(_rows(date(2026, 8, 12), PREFIX + "A"))
    store.save_snapshot(_rows(date(2026, 8, 19), PREFIX + "A"))
    dates = store.list_snapshot_dates()
    assert date(2026, 8, 12) in dates
    assert date(2026, 8, 19) in dates


@needs_db
def test_resaving_the_same_date_and_vendor_overwrites(store):
    store.save_snapshot(_rows(date(2026, 8, 12), PREFIX + "B", **{"Portal Files Found": 7}))
    store.save_snapshot(_rows(date(2026, 8, 12), PREFIX + "B", **{"Portal Files Found": 10}))
    hist = store.history(vendor_numbers=[PREFIX + "B"])
    assert len(hist) == 1
    assert hist[0].compliance_pct == pytest.approx(1.0)


@needs_db
def test_history_filters_by_vendor(store):
    store.save_snapshot(_rows(date(2026, 8, 12), PREFIX + "C"))
    store.save_snapshot(_rows(date(2026, 8, 12), PREFIX + "D"))
    hist = store.history(vendor_numbers=[PREFIX + "C"])
    assert [h.vendor_number for h in hist] == [PREFIX + "C"]


@needs_db
def test_history_orders_by_vendor_then_date(store):
    store.save_snapshot(_rows(date(2026, 8, 19), PREFIX + "E"))
    store.save_snapshot(_rows(date(2026, 8, 12), PREFIX + "E"))
    hist = store.history(vendor_numbers=[PREFIX + "E"])
    assert [h.snapshot_date for h in hist] == [date(2026, 8, 12), date(2026, 8, 19)]


@needs_db
def test_latest_two_dates_reflects_the_newest_snapshot(store):
    """Not asserted against an empty table (the shared test DB may carry
    other dates) — just that saving a later date surfaces as the newest."""
    store.save_snapshot(_rows(date(2026, 8, 12), PREFIX + "F"))
    store.save_snapshot(_rows(date(2099, 1, 1), PREFIX + "F"))
    latest = store.latest_two_dates()
    assert latest is not None
    assert latest[0] == date(2099, 1, 1)
