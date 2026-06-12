"""Regression tests for the compliance engine's portal/SAP matching."""
import pandas as pd

from src.compliance_engine import build_report


def _sap_df():
    """A single in-scope (June 2026) SAP PO with an inbound delivery."""
    return pd.DataFrame(
        {
            "PO Number": ["1001"],
            "Normalized PO Number": ["1001"],
            "Vendor Number": ["V1"],
            "Vendor Name": ["BOB'S RED MILL"],
            "Warehouse": ["W1"],
            "PO Status": ["A"],
            "Inbound Delivery": ["IB1"],
            "Inbound Delivery Status": ["A"],
            "Appointment Date": pd.to_datetime(["2026-06-02"]),
            "Delivery Date": pd.to_datetime(["2026-06-03"]),
            "Confirmed PU Date": pd.to_datetime(["2026-06-01"]),
            "Est PU Date": pd.to_datetime([pd.NaT]),
        }
    )


def _portal_df_no_invalid():
    """Portal export with ONLY an Approved upload -> invalid_lookup is empty.

    This reproduces the pandas 3.0 crash where mapping with an empty
    datetime-typed lookup Series raised 'Cannot cast DatetimeArray to float64'.
    """
    return pd.DataFrame(
        {
            "PO Number": ["1001"],
            "Normalized PO Number": ["1001"],
            "Supplier Name": ["BOB'S RED MILL"],
            "Upload Date": pd.to_datetime(["2026-06-04 10:00"]),
            "File Status": ["Approved"],
            "File Name": ["po1001.pdf"],
            "Uploaded By": ["bob@example.com"],
            "Downloaded By": [""],
            "Download Date": pd.to_datetime([pd.NaT]),
            "Invalid Comment": [""],
        }
    )


def test_build_report_with_no_invalid_uploads_does_not_crash():
    sheets = build_report(_sap_df(), _portal_df_no_invalid(), 2026, 6)
    matched = sheets["SAP Inbound Matched With Portal File"]
    assert len(matched) == 1
    # Upload Date must survive the matching (the lookup map must work).
    assert pd.notna(matched.iloc[0]["Upload Date"])


def test_build_report_with_no_portal_rows_does_not_crash():
    """Both valid_lookup and invalid_lookup empty -> no portal rows at all."""
    empty_portal = pd.DataFrame(
        {
            "PO Number": [],
            "Normalized PO Number": [],
            "Supplier Name": [],
            "Upload Date": pd.to_datetime([]),
            "File Status": [],
            "File Name": [],
            "Uploaded By": [],
            "Downloaded By": [],
            "Download Date": pd.to_datetime([]),
            "Invalid Comment": [],
        }
    )
    sheets = build_report(_sap_df(), empty_portal, 2026, 6)
    # PO has an inbound but no portal file -> shows up as missing, no crash.
    missing = sheets["SAP Inbound Missing Portal File"]
    assert len(missing) == 1
