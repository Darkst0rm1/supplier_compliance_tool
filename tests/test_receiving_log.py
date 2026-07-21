"""Tests for the Receiving Log import and the document-accuracy rollups.

The receiving log is hand-maintained, so most of these tests pin down how the
importer survives real-world mess: a banner above the header row, headers whose
spacing drifts, a mid-year schema change, and PO cells typed by hand.
"""
from io import BytesIO

import pandas as pd
import pytest

from src.compliance_engine import build_report
from src.receiving_importer import (
    ReceivingImportError,
    build_po_lookup,
    load_receiving,
    normalize_yes_no,
)


# ---------------------------------------------------------------------------
# Fixtures — workbooks shaped like the real file
# ---------------------------------------------------------------------------
def _write_book(sheets: dict[str, list[list]]) -> BytesIO:
    """Write raw rows (no header) to an in-memory xlsx, one sheet per key."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(
                writer, sheet_name=name, index=False, header=False
            )
    buf.seek(0)
    return buf


def _audit_sheet_rows():
    """The post-May schema: banner rows, then headers, then data.

    Header spacing is deliberately inconsistent ("Y / N" vs "Y/N") because the
    real workbook is inconsistent.
    """
    return [
        [None, "TOL - RECEIVING  LOG", None, None, None, None, None, None],
        [None, "Problem Key: 1-live or dead", None, None, None, None, None, None],
        [None, None, None, None, None, None, None, None],
        [
            "Date", "PO#", "Carrier", "Inbound File  Y / N",
            "Correct Batch Received Y / N", "Correct BBD Received Y/N",
            "Correct QTY Received Y/N", "Comments",
        ],
        ["2026-06-02", "1000001001", "NFI", "YES", "YES", "YES", "YES", None],
        ["2026-06-03", "1000001002", "CABUZZI", "NO", "NO", "YES", "YES", None],
        ["2026-06-04", "1000001003/1000001004", "PRECISION", "YES", "YES", "NO", "NO", None],
        ["2026-07-01", "1000001009", "NFI", "YES", "YES", "YES", "YES", None],
    ]


def _legacy_sheet_rows():
    """The Jan-Apr schema: no audit columns at all."""
    return [
        [None, None, None, "           RECEIVING  LOG", None],
        [None, None, None, None, None],
        [None, None, None, None, None],
        ["Date", "PO#", "Carrier", "Seal#", "Receiver Initials"],
        ["2026-06-05", "1000009999", "NFI", "1003798", "MS"],
    ]


def _receiving_book():
    return _write_book({"JAN": _legacy_sheet_rows(), "MAY - DEC": _audit_sheet_rows()})


# ---------------------------------------------------------------------------
# normalize_yes_no
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "raw,expected",
    [
        ("YES", "YES"), ("yes", "YES"), (" Yes ", "YES"), ("Y", "YES"),
        ("NO", "NO"), ("no", "NO"), ("N", "NO"),
        ("", ""), (None, ""), ("nan", ""), ("maybe", ""), ("N/A", ""),
    ],
)
def test_normalize_yes_no(raw, expected):
    assert normalize_yes_no(raw) == expected


def test_normalize_yes_no_handles_nan_float():
    assert normalize_yes_no(float("nan")) == ""


# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------
def test_finds_header_row_below_banner_and_parses_drifting_spacing():
    df = load_receiving(_receiving_book(), 2026, 6)
    # ...1001, ...1002, and ...1003 + ...1004 from the slash-separated cell.
    assert set(df["Normalized PO Number"]) == {
        "1000001001", "1000001002", "1000001003", "1000001004",
    }
    assert df.loc[df["Normalized PO Number"] == "1000001002", "Correct Batch"].iloc[0] == "NO"


def test_skips_sheets_without_audit_columns():
    """The legacy Jan-Apr schema must be skipped, not read as all-blank.

    Reading it would inject POs with no answers and dilute every rollup.
    """
    df = load_receiving(_receiving_book(), 2026, 6)
    assert "1000009999" not in set(df["Normalized PO Number"])
    assert "JAN" in df.attrs["skipped_sheets"]


def test_raises_when_no_sheet_has_audit_columns():
    book = _write_book({"JAN": _legacy_sheet_rows()})
    with pytest.raises(ReceivingImportError, match="audit columns"):
        load_receiving(book, 2026, 6)


def test_filters_to_report_month():
    """July's row must not leak into a June report."""
    df = load_receiving(_receiving_book(), 2026, 6)
    assert "1000001009" not in set(df["Normalized PO Number"])
    assert "1000001009" in set(load_receiving(_receiving_book(), 2026, 7)["Normalized PO Number"])


def test_month_with_no_rows_returns_empty_not_error():
    df = load_receiving(_receiving_book(), 2026, 1)
    assert df.empty
    assert build_po_lookup(df).empty


def test_splits_ampersand_separated_pos():
    rows = _audit_sheet_rows()
    rows[4][1] = "1000002001 & 1000002002"
    df = load_receiving(_write_book({"MAY - DEC": rows}), 2026, 6)
    assert {"1000002001", "1000002002"} <= set(df["Normalized PO Number"])


def test_drops_non_po_tokens_and_rows_without_a_po():
    """Hand-typed references like TR-34306 or "Return" are not SAP POs.

    They are kept out of the join rather than erroring, but a row with no
    usable PO at all is counted so the UI can report real coverage.
    """
    rows = _audit_sheet_rows()
    rows[4][1] = None
    rows[5][1] = "TR-34306"
    df = load_receiving(_write_book({"MAY - DEC": rows}), 2026, 6)
    assert df.attrs["rows_without_po"] == 1
    assert "TR-34306" not in set(df["Normalized PO Number"])


def test_excludes_six_series_pos():
    """The 6-series exclusion must apply to the log as it does to SAP/portal."""
    rows = _audit_sheet_rows()
    rows[4][1] = "6000002150"
    df = load_receiving(_write_book({"MAY - DEC": rows}), 2026, 6)
    assert "6000002150" not in set(df["Normalized PO Number"])
    assert df.attrs["excluded_po_count"] == 1


# ---------------------------------------------------------------------------
# build_po_lookup
# ---------------------------------------------------------------------------
def test_lookup_keeps_first_real_answer_not_first_row():
    """A blank row must not erase an answer recorded on another dock row.

    Split loads produce several rows for one PO, and the earlier one is often
    the one left unfilled.
    """
    rows = _audit_sheet_rows()
    rows[4] = ["2026-06-02", "1000001001", "NFI", "", "", "", "", None]
    rows[5] = ["2026-06-03", "1000001001", "NFI", "YES", "NO", "YES", "YES", None]
    df = load_receiving(_write_book({"MAY - DEC": rows}), 2026, 6)
    lookup = build_po_lookup(df)
    assert lookup.loc["1000001001", "Correct Batch"] == "NO"
    assert lookup.loc["1000001001", "Inbound File Received"] == "YES"


# ---------------------------------------------------------------------------
# Engine integration
# ---------------------------------------------------------------------------
def _sap_df():
    """Four June 2026 POs with inbound deliveries, two suppliers."""
    return pd.DataFrame(
        {
            "PO Number": ["1000001001", "1000001002", "1000001003", "1000001005"],
            "Normalized PO Number": ["1000001001", "1000001002", "1000001003", "1000001005"],
            "Vendor Number": ["V1", "V1", "V2", "V2"],
            "Vendor Name": ["ACME", "ACME", "BETA", "BETA"],
            "Warehouse": ["W1", "W1", "W1", "W1"],
            "PO Status": ["A", "A", "A", "A"],
            "Inbound Delivery": ["IB1", "IB2", "IB3", "IB5"],
            "Inbound Delivery Status": ["A", "A", "A", "A"],
            "Appointment Date": pd.to_datetime(["2026-06-02"] * 4),
            "Delivery Date": pd.to_datetime(["2026-06-03"] * 4),
            "Confirmed PU Date": pd.to_datetime([pd.NaT] * 4),
            "Est PU Date": pd.to_datetime([pd.NaT] * 4),
        }
    )


def _portal_df(pos):
    return pd.DataFrame(
        {
            "PO Number": pos,
            "Normalized PO Number": pos,
            "Supplier Name": ["ACME"] * len(pos),
            "Upload Date": pd.to_datetime(["2026-06-04 10:00"] * len(pos)),
            "File Status": ["Approved"] * len(pos),
            "File Name": ["f.pdf"] * len(pos),
            "Uploaded By": ["a@b.c"] * len(pos),
            "Downloaded By": [""] * len(pos),
            "Download Date": pd.to_datetime([pd.NaT] * len(pos)),
            "Invalid Comment": [""] * len(pos),
        }
    )


def test_receiving_log_is_optional_and_omits_its_sheets():
    sheets = build_report(_sap_df(), _portal_df(["1000001001"]), 2026, 6)
    assert "Document Accuracy Exceptions" not in sheets
    assert "Doc Accuracy Checked POs" not in sheets["Supplier Summary"].columns


def test_document_accuracy_never_changes_compliance_or_billback():
    """The audited number must mean exactly what it meant before.

    Accuracy is a separate dimension; folding a wrong BBD into the compliance
    percentage would silently redefine a figure people sign off on.
    """
    sap, portal = _sap_df(), _portal_df(["1000001001"])
    without = build_report(sap, portal, 2026, 6)
    with_log = build_report(
        sap, portal, 2026, 6, receiving_df=load_receiving(_receiving_book(), 2026, 6)
    )

    def _metric(sheets, key):
        return sheets["Monthly Summary"].set_index("Metric")["Value"][key]

    assert _metric(without, "Compliance Percentage") == _metric(with_log, "Compliance Percentage")
    assert (
        _metric(without, "SAP Inbound POs Missing Portal File")
        == _metric(with_log, "SAP Inbound POs Missing Portal File")
    )
    assert {k for k in without if k.startswith("BB-")} == {
        k for k in with_log if k.startswith("BB-")
    }


def test_accuracy_denominator_counts_only_checked_pos():
    """Unanswered POs must not be scored as passes.

    Most of the log is blank; counting blanks as YES would manufacture a high
    accuracy rate that nobody actually measured.
    """
    sheets = build_report(
        _sap_df(), _portal_df(["1000001001"]), 2026, 6,
        receiving_df=load_receiving(_receiving_book(), 2026, 6),
    )
    summary = sheets["Monthly Summary"].set_index("Metric")["Value"]
    # 1001, 1002, 1003 are checked in the log; 1005 is absent from it.
    assert summary["SAP POs With A Document Accuracy Check"] == 3
    assert summary["POs Failing Any Accuracy Check"] == 2  # 1002 batch, 1003 BBD+QTY
    assert summary["Document Accuracy Percentage"] == "33.3%"

    supplier = sheets["Supplier Summary"].set_index("Vendor Name")
    assert supplier.loc["BETA", "Doc Accuracy Checked POs"] == 1
    assert supplier.loc["BETA", "Wrong QTY"] == 1


def test_supplier_with_no_checks_shows_na_not_zero_percent():
    """A supplier the dock never checked is unmeasured, not perfect or failing."""
    rows = _audit_sheet_rows()
    del rows[6]  # drop the row covering BETA's PO 1003
    sheets = build_report(
        _sap_df(), _portal_df(["1000001001"]), 2026, 6,
        receiving_df=load_receiving(_write_book({"MAY - DEC": rows}), 2026, 6),
    )
    supplier = sheets["Supplier Summary"].set_index("Vendor Name")
    assert supplier.loc["BETA", "Document Accuracy Percentage"] == "n/a"


def test_accuracy_exceptions_name_the_failing_checks():
    sheets = build_report(
        _sap_df(), _portal_df(["1000001001"]), 2026, 6,
        receiving_df=load_receiving(_receiving_book(), 2026, 6),
    )
    exc = sheets["Document Accuracy Exceptions"].set_index("PO Number")
    assert "batch" in exc.loc["1000001002", "Issue"]
    assert exc.loc["1000001003", "Issue"] == (
        "Document did not match goods received: BBD, quantity."
    )


def test_flags_portal_and_dock_disagreements_both_ways():
    """1002: dock saw no file, portal has one. 1001: both agree -> not flagged."""
    sheets = build_report(
        _sap_df(), _portal_df(["1000001001", "1000001002"]), 2026, 6,
        receiving_df=load_receiving(_receiving_book(), 2026, 6),
    )
    conflicts = sheets["Portal vs Receiving Log"].set_index("PO Number")
    assert "1000001001" not in conflicts.index
    assert conflicts.loc["1000001002", "Disagreement"].startswith("Portal has a file")
    # 1003: dock recorded a file received but no portal upload exists.
    assert conflicts.loc["1000001003", "Disagreement"].startswith("Dock recorded")
