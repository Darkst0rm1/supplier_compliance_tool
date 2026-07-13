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
