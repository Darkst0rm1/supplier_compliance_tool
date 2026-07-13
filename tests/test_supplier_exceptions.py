"""Tests for the supplier exceptions feature."""
from __future__ import annotations

import pytest

from src.normalizer import normalize_supplier_name


class TestNormalizeSupplierName:
    def test_uppercases_and_strips(self):
        assert normalize_supplier_name("  Acetum S.P.A.  ") == "ACETUM S P A"

    def test_collapses_internal_whitespace(self):
        assert normalize_supplier_name("BOTHWELL    CHEESE") == "BOTHWELL CHEESE"

    def test_punctuation_becomes_space(self):
        # The tracker and SAP punctuate differently; both must land on the same key.
        assert normalize_supplier_name("DARE (LESLEY STOWE FINE FOODS)") == (
            "DARE LESLEY STOWE FINE FOODS"
        )
        assert normalize_supplier_name("D&D ITALIA SPA") == "D D ITALIA SPA"
        assert normalize_supplier_name("C.H. GUENTHER & SON, INC.") == (
            "C H GUENTHER SON INC"
        )

    def test_sap_and_tracker_spellings_converge(self):
        # Real pair from the June export vs the tracker.
        assert normalize_supplier_name("SERFUNGHI DI CALABRETTA LUIGI") == (
            "SERFUNGHI DI CALABRETTA LUIGI"
        )
        assert normalize_supplier_name("Serfunghi di Calabretta Luigi") == (
            "SERFUNGHI DI CALABRETTA LUIGI"
        )

    def test_empty_and_missing(self):
        assert normalize_supplier_name(None) == ""
        assert normalize_supplier_name("") == ""
        assert normalize_supplier_name("   ") == ""
        assert normalize_supplier_name(float("nan")) == ""

    def test_curly_apostrophe_converges_with_straight(self):
        # Excel/Word autocorrect turns ' into the curly U+2019 variant.
        assert normalize_supplier_name("O’Brien Foods") == (
            normalize_supplier_name("O'Brien Foods")
        )
        assert normalize_supplier_name("O'Brien Foods") == "O BRIEN FOODS"

    def test_accented_name_converges_with_unaccented(self):
        # The supplier set is heavily Italian/European.
        assert normalize_supplier_name("Caffè Mauro") == (
            normalize_supplier_name("CAFFE MAURO")
        )
        assert normalize_supplier_name("Caffè Mauro") == "CAFFE MAURO"

    def test_em_dash_becomes_space(self):
        assert normalize_supplier_name("Acme—Foods") == "ACME FOODS"
        assert normalize_supplier_name("Acme–Foods") == "ACME FOODS"


import pandas as pd

from src.config import REASON_EXEMPT_MARK, REASON_UNABLE_TO_COMPLY
from src.tracker_importer import TrackerImportError, read_tracker_exceptions


def _fake_tracker(tmp_path):
    """A miniature stand-in for the real tracker workbook."""
    path = tmp_path / "tracker.xlsx"
    tracker = pd.DataFrame({
        "Supplier Names ": [
            "ACETUM S.P.A.",
            "BOTHWELL CHEESE",
            "DARE (LESLEY STOWE FINE FOODS)",
            "COMPLIANT CO",
        ],
        "Compliance Status": [
            "NO -  Unable to Comply",   # double space, as in the real workbook
            "NO -  Unable to Comply",
            "NO -  Unable to Comply",
            "YES - Submitted on Portal",
        ],
    })
    # "POs received": name in col 0, the EXEMPT marker in col 3.
    pos_received = pd.DataFrame({
        "Non compliant": ["BOTHWELL CHEESE", "LUNDBERG FAMILY FARMS", "OTHER CO"],
        "Pos received": [15, 12, 3],
        "Column1": ["x", "x", "x"],
        "Unnamed: 3": ["EXEMPT", "EXEMPT", None],
    })
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        tracker.to_excel(w, sheet_name="Tracker", index=False)
        pos_received.to_excel(w, sheet_name="POs received", index=False)
    return path


class TestReadTrackerExceptions:
    def test_unions_both_lists_and_dedupes(self, tmp_path):
        rows = read_tracker_exceptions(_fake_tracker(tmp_path))
        names = {n for n, _ in rows}
        # 3 Unable-to-Comply + 2 EXEMPT-marked, overlapping on Bothwell -> 4.
        assert names == {
            "ACETUM S.P.A.",
            "BOTHWELL CHEESE",
            "DARE (LESLEY STOWE FINE FOODS)",
            "LUNDBERG FAMILY FARMS",
        }
        assert "COMPLIANT CO" not in names
        assert "OTHER CO" not in names

    def test_unable_to_comply_wins_the_reason_on_overlap(self, tmp_path):
        rows = dict(read_tracker_exceptions(_fake_tracker(tmp_path)))
        assert rows["BOTHWELL CHEESE"] == REASON_UNABLE_TO_COMPLY
        assert rows["LUNDBERG FAMILY FARMS"] == REASON_EXEMPT_MARK

    def test_missing_sheet_raises_friendly_error(self, tmp_path):
        path = tmp_path / "wrong.xlsx"
        pd.DataFrame({"a": [1]}).to_excel(path, sheet_name="Nope", index=False)
        with pytest.raises(TrackerImportError, match="Tracker"):
            read_tracker_exceptions(path)
