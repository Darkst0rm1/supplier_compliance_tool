from src.compliance_engine import _billback_sheet_name


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


import pandas as pd

from src.compliance_engine import _billback_supplier_tab
from src.config import BILLBACK_FEE_PER_OCCURRENCE, BILLBACK_REASON


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
