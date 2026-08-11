"""Tests for the Supplier Price Summary sheet and the Supplier Summary BDM columns."""
from __future__ import annotations

import pandas as pd
import pytest

from src.compliance_engine import build_report
from src.config import BILLBACK_FEE_PER_OCCURRENCE
from src.normalizer import normalize_supplier_name
from src.supplier_exceptions import ExceptionRecord


def _sap_row(po, vendor_num, vendor_name, inbound="IBD-1", bdm="", bdm_description=""):
    return {
        "PO Number": po,
        "Normalized PO Number": po,
        "Vendor Number": vendor_num,
        "Vendor Name": vendor_name,
        "Warehouse": "WH1",
        "PO Status": "A",
        "Appointment Date": pd.Timestamp("2026-06-15"),
        "Delivery Date": pd.Timestamp("2026-06-15"),
        "Confirmed PU Date": pd.NaT,
        "Est PU Date": pd.NaT,
        "Inbound Delivery": inbound,
        "Inbound Delivery Status": "A",
        "BDM": bdm,
        "BDM Description": bdm_description,
    }


def _portal_row(po, supplier, status="Approved"):
    return {
        "PO Number": po,
        "Normalized PO Number": po,
        "Supplier Name": supplier,
        "Upload Date": pd.Timestamp("2026-06-16"),
        "File Status": status,
        "File Name": "doc.pdf",
        "Uploaded By": "someone",
        "Invalid Comment": "",
        "Downloaded By": "",
        "Download Date": pd.NaT,
    }


def _rec(name):
    key = normalize_supplier_name(name)
    return ExceptionRecord(
        id=1, supplier_name=name, normalized_name=key,
        vendor_number=None, reason="Manual",
    )


@pytest.fixture
def scenario():
    """Three vendors, one PO each unless noted:
      COMPLIANT CO   -- uploaded its file            -> $0, Exempt No
      MISSING DOCS   -- two POs, uploaded neither     -> $400, Exempt No
      EXEMPT CO      -- an approved exception, uploaded nothing -> still $200
                         (exceptions are informational only)
    """
    sap = pd.DataFrame([
        _sap_row("1001", "70001111", "COMPLIANT CO"),
        _sap_row("1002", "70002222", "MISSING DOCS"),
        _sap_row("1003", "70002222", "MISSING DOCS"),
        _sap_row("1004", "70003333", "EXEMPT CO"),
    ])
    portal = pd.DataFrame([_portal_row("1001", "COMPLIANT CO")])
    exceptions = {normalize_supplier_name("EXEMPT CO"): _rec("EXEMPT CO")}
    return sap, portal, exceptions


class TestSupplierPriceSummary:
    def test_sheet_present_with_expected_columns(self, scenario):
        sap, portal, exceptions = scenario
        sheets = build_report(sap, portal, 2026, 6, exceptions=exceptions)
        assert list(sheets["Supplier Price Summary"].columns) == [
            "Vendor Number", "Vendor Name", "Exempt",
            "Missing Documents Billed", "Price (USD)",
        ]

    def test_compliant_vendor_is_zero(self, scenario):
        sap, portal, exceptions = scenario
        sheets = build_report(sap, portal, 2026, 6, exceptions=exceptions)
        row = sheets["Supplier Price Summary"].set_index("Vendor Name").loc["COMPLIANT CO"]
        assert row["Price (USD)"] == 0
        assert row["Missing Documents Billed"] == 0
        assert row["Exempt"] == "No"

    def test_missing_docs_charged_per_occurrence(self, scenario):
        sap, portal, exceptions = scenario
        sheets = build_report(sap, portal, 2026, 6, exceptions=exceptions)
        row = sheets["Supplier Price Summary"].set_index("Vendor Name").loc["MISSING DOCS"]
        assert row["Missing Documents Billed"] == 2
        assert row["Price (USD)"] == 2 * BILLBACK_FEE_PER_OCCURRENCE

    def test_exempt_vendor_still_billed_but_flagged(self, scenario):
        # Exceptions are informational-only: the flag says Yes, but the
        # charge does not change. See TestBillbackAndComplianceUnchanged in
        # test_supplier_exceptions.py for the same invariant on the BB- tabs.
        sap, portal, exceptions = scenario
        sheets = build_report(sap, portal, 2026, 6, exceptions=exceptions)
        row = sheets["Supplier Price Summary"].set_index("Vendor Name").loc["EXEMPT CO"]
        assert row["Exempt"] == "Yes"
        assert row["Price (USD)"] == BILLBACK_FEE_PER_OCCURRENCE

    def test_every_vendor_in_scope_gets_a_row(self, scenario):
        sap, portal, exceptions = scenario
        sheets = build_report(sap, portal, 2026, 6, exceptions=exceptions)
        names = set(sheets["Supplier Price Summary"]["Vendor Name"])
        assert names == {"COMPLIANT CO", "MISSING DOCS", "EXEMPT CO"}

    def test_sorted_by_price_descending(self, scenario):
        sap, portal, exceptions = scenario
        sheets = build_report(sap, portal, 2026, 6, exceptions=exceptions)
        prices = sheets["Supplier Price Summary"]["Price (USD)"].tolist()
        assert prices == sorted(prices, reverse=True)

    def test_invalid_upload_excluded_from_price_like_billback(self):
        # Mirrors test_billsheets_excludes_invalid_uploads in test_billback.py.
        sap = pd.DataFrame([_sap_row("9001", "70009999", "REJECTED UPLOAD")])
        portal = pd.DataFrame([_portal_row("9001", "REJECTED UPLOAD", status="Invalid")])
        sheets = build_report(sap, portal, 2026, 6)
        row = sheets["Supplier Price Summary"].set_index("Vendor Name").loc["REJECTED UPLOAD"]
        assert row["Price (USD)"] == 0
        assert row["Missing Documents Billed"] == 0


class TestSupplierSummaryBdmColumns:
    """BDM/BDM Description come straight from the SAP export's own columns --
    not every SAP vintage carries them, so absence just means blanks."""

    def test_columns_present_and_blank_when_sap_has_no_bdm(self, scenario):
        sap, portal, exceptions = scenario
        sheets = build_report(sap, portal, 2026, 6, exceptions=exceptions)
        summary = sheets["Supplier Summary"]
        assert "BDM" in summary.columns
        assert "BDM Description" in summary.columns
        assert set(summary["BDM"]) == {""}

    def test_vendor_with_bdm_on_its_sap_rows_gets_it(self):
        sap = pd.DataFrame([
            _sap_row("2001", "70007212", "ACQUA MINERALE", bdm="G6", bdm_description="ASIF AHMAD"),
        ])
        portal = pd.DataFrame(columns=list(_portal_row("x", "y")))
        sheets = build_report(sap, portal, 2026, 6)
        row = sheets["Supplier Summary"].set_index("Vendor Name").loc["ACQUA MINERALE"]
        assert row["BDM"] == "G6"
        assert row["BDM Description"] == "ASIF AHMAD"

    def test_vendor_with_no_bdm_on_its_sap_rows_is_blank(self, scenario):
        sap, portal, exceptions = scenario
        sheets = build_report(sap, portal, 2026, 6, exceptions=exceptions)
        row = sheets["Supplier Summary"].set_index("Vendor Name").loc["MISSING DOCS"]
        assert row["BDM"] == ""
        assert row["BDM Description"] == ""

    def test_a_blank_bdm_on_one_po_does_not_hide_a_real_one_on_another(self):
        # Same vendor, two POs: the first row's BDM is blank (a data-entry
        # gap), the second carries the real value. The blank first row must
        # not win.
        sap = pd.DataFrame([
            _sap_row("3001", "70009999", "TWO PO VENDOR", bdm="", bdm_description=""),
            _sap_row("3002", "70009999", "TWO PO VENDOR", bdm="K1", bdm_description="KARISHMA SALIAN"),
        ])
        portal = pd.DataFrame(columns=list(_portal_row("x", "y")))
        sheets = build_report(sap, portal, 2026, 6)
        row = sheets["Supplier Summary"].set_index("Vendor Name").loc["TWO PO VENDOR"]
        assert row["BDM"] == "K1"
        assert row["BDM Description"] == "KARISHMA SALIAN"
