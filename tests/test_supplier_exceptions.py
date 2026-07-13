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


from src.config import (
    EXCEPTION_STATUS_EXCEPTION,
    EXCEPTION_STATUS_EXPECTED,
    EXCEPTION_STATUS_NOT_ON_TRACKER,
)
from src.supplier_exceptions import (
    ExceptionRecord,
    ExceptionValidationError,
    classify_supplier,
    validate_supplier_name,
)


def _rec(name, reason="Unable to Comply", vendor_number=None):
    return ExceptionRecord(
        id=1,
        supplier_name=name,
        normalized_name=normalize_supplier_name(name),
        vendor_number=vendor_number,
        reason=reason,
        added_by=None,
        added_at=None,
    )


class TestClassifySupplier:
    def setup_method(self):
        self.exceptions = {
            normalize_supplier_name("BOTHWELL CHEESE"): _rec("BOTHWELL CHEESE"),
            normalize_supplier_name("CAFFE MAURO SPA"): _rec(
                "CAFFE MAURO SPA", vendor_number="70006979"
            ),
        }
        self.tracker = {
            normalize_supplier_name("BOTHWELL CHEESE"),
            normalize_supplier_name("CAFFE MAURO SPA"),
            normalize_supplier_name("ACQUA MINERALE SAN BENEDETTO"),
        }

    def test_exception_supplier(self):
        assert classify_supplier(
            "Bothwell Cheese", "70001111", self.exceptions, self.tracker
        ) == EXCEPTION_STATUS_EXCEPTION

    def test_vendor_number_matches_even_when_the_name_differs(self):
        # SAP spells it differently, but we recorded the vendor number.
        assert classify_supplier(
            "CAFFE MAURO S.P.A. (ITALY)", "70006979", self.exceptions, self.tracker
        ) == EXCEPTION_STATUS_EXCEPTION

    def test_on_tracker_but_not_an_exception(self):
        assert classify_supplier(
            "ACQUA MINERALE SAN BENEDETTO", "70007212", self.exceptions, self.tracker
        ) == EXCEPTION_STATUS_EXPECTED

    def test_absent_from_the_tracker(self):
        # A 3PL warehouse, not a supplier at all.
        assert classify_supplier(
            "AMERICOLD TACOMA", "70009999", self.exceptions, self.tracker
        ) == EXCEPTION_STATUS_NOT_ON_TRACKER

    def test_empty_tracker_collapses_to_expected(self):
        assert classify_supplier(
            "ANYONE", "70009999", self.exceptions, set()
        ) == EXCEPTION_STATUS_EXPECTED

    def test_blank_vendor_name(self):
        assert classify_supplier("", "", self.exceptions, self.tracker) == (
            EXCEPTION_STATUS_NOT_ON_TRACKER
        )


class TestValidateSupplierName:
    def test_rejects_blank(self):
        with pytest.raises(ExceptionValidationError):
            validate_supplier_name("   ")

    def test_trims(self):
        assert validate_supplier_name("  Acme  ") == "Acme"


import os

from src.config import REASON_MANUAL
from src.supplier_exceptions import (
    DuplicateExceptionError,
    ExceptionNotFoundError,
    ExceptionStore,
)

_TEST_DSN = os.environ.get("TEST_DATABASE_URL")
requires_db = pytest.mark.skipif(not _TEST_DSN, reason="TEST_DATABASE_URL not set")


@requires_db
class TestExceptionStore:
    def setup_method(self):
        self.store = ExceptionStore(_TEST_DSN)
        self.store.ensure_schema()
        self.name = "PYTEST TEMP SUPPLIER"
        try:
            self.store.remove_exception(normalize_supplier_name(self.name))
        except ExceptionNotFoundError:
            pass

    def teardown_method(self):
        try:
            self.store.remove_exception(normalize_supplier_name(self.name))
        except ExceptionNotFoundError:
            pass

    def test_add_load_remove_roundtrip(self):
        rec = self.store.add_exception(self.name, REASON_MANUAL)
        assert rec.normalized_name == normalize_supplier_name(self.name)

        loaded = self.store.load_exceptions()
        assert rec.normalized_name in loaded
        assert loaded[rec.normalized_name].reason == REASON_MANUAL

        self.store.remove_exception(rec.normalized_name)
        assert rec.normalized_name not in self.store.load_exceptions()

    def test_duplicate_rejected(self):
        self.store.add_exception(self.name, REASON_MANUAL)
        with pytest.raises(DuplicateExceptionError):
            self.store.add_exception(self.name.lower(), REASON_MANUAL)

    def test_remove_missing_raises(self):
        with pytest.raises(ExceptionNotFoundError):
            self.store.remove_exception("NO SUCH SUPPLIER AT ALL")
