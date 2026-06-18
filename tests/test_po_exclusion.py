"""POs starting with an excluded prefix (e.g. '6') are disregarded on import."""
from __future__ import annotations

import io

import pandas as pd

from src.normalizer import is_excluded_po
from src.portal_importer import load_portal
from src.sap_importer import load_sap


def _xlsx(df: pd.DataFrame) -> io.BytesIO:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


def test_is_excluded_po():
    assert is_excluded_po("6000001873", ("6",))
    assert not is_excluded_po("1000007174", ("6",))
    assert not is_excluded_po("6000001873", ())   # no prefixes -> nothing excluded
    assert not is_excluded_po("", ("6",))


def test_sap_excludes_six_series():
    sap = pd.DataFrame({
        "PO Number": ["1000004001", "6000001873", "1000004715", "6000002005"],
        "Inbound Delivery": ["180013022", "180013023", "180012747", "180012748"],
        "Inbound Delivery Status": ["C", "C", "C", "C"],
    })
    out = load_sap(_xlsx(sap))
    pos = set(out["Normalized PO Number"])
    assert pos == {"1000004001", "1000004715"}
    assert not any(p.startswith("6") for p in pos)
    assert out.attrs["excluded_po_count"] == 2


def test_portal_excludes_six_series():
    portal = pd.DataFrame({
        "PO Number(s)": ["1000007174", "6000001999", "1000007162,6000002001"],
        "Supplier": ["KIKKOMAN", "KIKKOMAN", "DEVANCO"],
        "Upload Date": ["6/17/2026, 7:31 PM", "6/17/2026, 6:45 PM", "6/17/2026, 6:38 PM"],
    })
    out = load_portal(_xlsx(portal), 2026, 6)
    pos = set(out["Normalized PO Number"])
    assert pos == {"1000007174", "1000007162"}
    assert not any(p.startswith("6") for p in pos)
    # two 6-series POs across the rows (one standalone, one inside a multi-PO cell)
    assert out.attrs["excluded_po_count"] == 2
