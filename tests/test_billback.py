import pandas as pd

from src.compliance_engine import (
    _billback_sheet_name,
    _billback_sheets,
    _billback_supplier_tab,
)
from src.config import BILLBACK_FEE_PER_OCCURRENCE, BILLBACK_REASON


def test_sheet_name_prefixes_and_keeps_simple_name():
    used = set()
    assert _billback_sheet_name("BOB'S RED MILL", "1001", used) == "BB-BOB'S RED MILL"


def test_sheet_name_strips_illegal_chars():
    used = set()
    name = _billback_sheet_name("A/B:C*D[E]F?G\\H", "1", used)
    for ch in r":\/?*[]":
        assert ch not in name
    assert name.startswith("BB-")


def test_sheet_name_truncates_to_31_chars():
    used = set()
    name = _billback_sheet_name("X" * 60, "1", used)
    assert len(name) <= 31
    assert name.startswith("BB-")


def test_sheet_name_dedupes_collisions():
    used = set()
    first = _billback_sheet_name("SAME NAME", "1", used)
    second = _billback_sheet_name("SAME NAME", "2", used)
    assert first != second
    assert first in used and second in used
    assert len(second) <= 31


def test_sheet_name_falls_back_to_number_then_unknown():
    used = set()
    assert _billback_sheet_name("", "9999", used) == "BB-9999"
    assert _billback_sheet_name("", "", used) == "BB-Unknown Supplier"


def test_sheet_name_falls_back_when_name_is_all_illegal_chars():
    used = set()
    # vendor_name is non-empty but all illegal chars -> must use vendor_number
    assert _billback_sheet_name("///", "9999", used) == "BB-9999"


def _sample_missing_rows():
    return pd.DataFrame(
        {
            "PO Number": ["1001", "1002"],
            "Warehouse": ["W1", "W1"],
            "PO Status": ["B", "B"],
            "Appointment Date": ["2026-06-01", "2026-06-02"],
            "Delivery Date": ["2026-06-03", "2026-06-04"],
            "Inbound Delivery": ["IB1", "IB2"],
        }
    )


def test_supplier_tab_has_expected_columns():
    tab = _billback_supplier_tab(_sample_missing_rows())
    assert list(tab.columns) == [
        "PO Number", "Warehouse", "PO Status", "Appointment Date",
        "Delivery Date", "Inbound Delivery", "Charge Reason", "Charge (USD)",
    ]


def test_supplier_tab_charges_fee_per_po():
    tab = _billback_supplier_tab(_sample_missing_rows())
    # First 2 rows are POs, last row is the TOTAL row.
    po_rows = tab.iloc[:-1]
    assert (po_rows["Charge (USD)"] == BILLBACK_FEE_PER_OCCURRENCE).all()
    assert (po_rows["Charge Reason"] == BILLBACK_REASON).all()


def test_supplier_tab_total_row():
    tab = _billback_supplier_tab(_sample_missing_rows())
    total = tab.iloc[-1]
    assert "2 occurrences" in str(total["PO Number"])
    assert total["Charge (USD)"] == 2 * BILLBACK_FEE_PER_OCCURRENCE


def _missing_bucket():
    # Two never-uploaded POs for vendor 1001, one for vendor 2002,
    # plus one Invalid-upload PO that must be excluded from billing.
    return pd.DataFrame(
        {
            "PO Number": ["1001", "1002", "2001", "9001"],
            "Vendor Number": ["1001", "1001", "2002", "1001"],
            "Vendor Name": ["BOB'S RED MILL", "BOB'S RED MILL",
                            "HP HOOD LLC", "BOB'S RED MILL"],
            "Warehouse": ["W1", "W1", "W2", "W1"],
            "PO Status": ["B", "B", "B", "B"],
            "Appointment Date": ["2026-06-01"] * 4,
            "Delivery Date": ["2026-06-03"] * 4,
            "Inbound Delivery": ["IB1", "IB2", "IB3", "IB9"],
            "Portal Invalid Match": [False, False, False, True],
        }
    )


def test_billsheets_excludes_invalid_uploads():
    sheets = _billback_sheets(_missing_bucket())
    bob_name = next(n for n in sheets if "BOB" in n)
    bob_tab = sheets[bob_name]
    po_rows = bob_tab.iloc[:-1]
    assert set(po_rows["PO Number"]) == {"1001", "1002"}
    assert "9001" not in set(po_rows["PO Number"])


def test_billsheets_one_tab_per_supplier():
    sheets = _billback_sheets(_missing_bucket())
    assert len(sheets) == 2  # BOB'S RED MILL + HP HOOD LLC


def test_billsheets_ordered_by_occurrences_desc():
    sheets = _billback_sheets(_missing_bucket())
    first_name = next(iter(sheets))  # dict preserves insertion order
    assert "BOB" in first_name  # 2 occurrences ranks above HP HOOD's 1


def test_billsheets_empty_when_no_billable():
    empty = _missing_bucket().iloc[0:0]
    assert _billback_sheets(empty) == {}


def test_billsheets_empty_when_all_invalid():
    df = _missing_bucket().copy()
    df["Portal Invalid Match"] = True
    assert _billback_sheets(df) == {}
