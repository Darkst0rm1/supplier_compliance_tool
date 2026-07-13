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

    def test_works_from_an_in_memory_buffer_not_just_a_path(self, tmp_path):
        # Regression: _unable_to_comply and _exempt_marked both read the same
        # object. A Streamlit UploadedFile (unlike a path) is exhausted by the
        # first pd.read_excel call unless each read seeks back to 0 first --
        # without that, the EXEMPT-marked rows would silently vanish.
        import io

        buffer = io.BytesIO(_fake_tracker(tmp_path).read_bytes())
        rows = read_tracker_exceptions(buffer)
        names = {n for n, _ in rows}
        assert names == {
            "ACETUM S.P.A.",
            "BOTHWELL CHEESE",
            "DARE (LESLEY STOWE FINE FOODS)",
            "LUNDBERG FAMILY FARMS",
        }


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


class _FakeConn:
    """Stand-in for a psycopg connection: supports `with conn:` and
    `conn.execute(...).fetchall()`, returning canned rows."""

    def __init__(self, rows):
        self._rows = rows

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        return False

    def execute(self, *args, **kwargs):
        return self

    def fetchall(self):
        return self._rows


class TestLoadExceptionsSelfHeals:
    """Finding 1: load_exceptions() must not trust the stored normalized_name
    column -- it must recompute the key from supplier_name so classify_supplier
    (which always recomputes) still matches even if the normalizer changed
    after the row was written."""

    def test_stale_normalized_name_column_is_ignored(self, monkeypatch):
        stale_row = {
            "id": 1,
            "supplier_name": "Caffè Mauro",
            # Deliberately wrong/stale relative to the current normalizer --
            # as if this row was written before diacritic folding existed.
            "normalized_name": "CAFF MAURO OLD STALE VALUE",
            "vendor_number": None,
            "reason": REASON_MANUAL,
            "added_by": None,
            "added_at": None,
        }
        store = ExceptionStore("postgresql://unused/dsn")
        monkeypatch.setattr(store, "_connect", lambda: _FakeConn([stale_row]))

        loaded = store.load_exceptions()

        correct_key = normalize_supplier_name("Caffè Mauro")
        assert correct_key in loaded
        assert "CAFF MAURO OLD STALE VALUE" not in loaded
        # The record itself is internally consistent too.
        assert loaded[correct_key].normalized_name == correct_key

        # And classify_supplier -- which recomputes the key independently --
        # actually matches against it.
        assert classify_supplier("Caffe Mauro", None, loaded, set()) == (
            EXCEPTION_STATUS_EXCEPTION
        )


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


from src.compliance_engine import build_report


def _sap_row(po, vendor_num, vendor_name, inbound="IBD-1"):
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


@pytest.fixture
def scenario():
    """Three suppliers:
      BOTHWELL CHEESE  -- an exception, uploaded nothing
      ACQUA MINERALE   -- on the tracker, uploaded nothing  <- should be chased
      AMERICOLD TACOMA -- not on the tracker, uploaded nothing
      GOOD SUPPLIER    -- on the tracker, uploaded its file
    """
    sap = pd.DataFrame([
        _sap_row("1001", "70001111", "BOTHWELL CHEESE"),
        _sap_row("1002", "70007212", "ACQUA MINERALE SAN BENEDETTO"),
        _sap_row("1003", "70007212", "ACQUA MINERALE SAN BENEDETTO"),
        _sap_row("1004", "70009999", "AMERICOLD TACOMA"),
        _sap_row("1005", "70002222", "GOOD SUPPLIER"),
    ])
    portal = pd.DataFrame([_portal_row("1005", "GOOD SUPPLIER")])
    exceptions = {
        normalize_supplier_name("BOTHWELL CHEESE"): _rec("BOTHWELL CHEESE"),
    }
    tracker = {
        normalize_supplier_name("BOTHWELL CHEESE"),
        normalize_supplier_name("ACQUA MINERALE SAN BENEDETTO"),
        normalize_supplier_name("GOOD SUPPLIER"),
    }
    return sap, portal, exceptions, tracker


class TestSupplierSummaryExceptionColumn:
    def test_column_present_with_the_three_states(self, scenario):
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        summary = sheets["Supplier Summary"].set_index("Vendor Name")

        assert "Exception Status" in sheets["Supplier Summary"].columns
        assert summary.loc["BOTHWELL CHEESE", "Exception Status"] == (
            EXCEPTION_STATUS_EXCEPTION
        )
        assert summary.loc["ACQUA MINERALE SAN BENEDETTO", "Exception Status"] == (
            EXCEPTION_STATUS_EXPECTED
        )
        assert summary.loc["AMERICOLD TACOMA", "Exception Status"] == (
            EXCEPTION_STATUS_NOT_ON_TRACKER
        )

    def test_without_exceptions_column_still_exists(self, scenario):
        # No DB / no exceptions passed: the column reads "Expected to upload"
        # for everyone rather than vanishing, so the sheet's shape is stable.
        sap, portal, _, _ = scenario
        sheets = build_report(sap, portal, 2026, 6)
        col = sheets["Supplier Summary"]["Exception Status"]
        assert set(col) == {EXCEPTION_STATUS_EXPECTED}


class TestBillbackAndComplianceUnchanged:
    """The load-bearing guarantee: exceptions are INFORMATIONAL ONLY."""

    def test_billback_identical_with_and_without_exceptions(self, scenario):
        sap, portal, exceptions, tracker = scenario
        without = build_report(sap, portal, 2026, 6)
        with_exc = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )

        bb_without = {k: v for k, v in without.items() if k.startswith("BB-")}
        bb_with = {k: v for k, v in with_exc.items() if k.startswith("BB-")}

        assert set(bb_without) == set(bb_with)
        # Bothwell is an exception but is STILL billed -- by design, for now.
        assert any("BOTHWELL" in k.upper() for k in bb_with)
        for name in bb_without:
            pd.testing.assert_frame_equal(bb_without[name], bb_with[name])

    def test_compliance_percentage_identical(self, scenario):
        sap, portal, exceptions, tracker = scenario
        without = build_report(sap, portal, 2026, 6)
        with_exc = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        pd.testing.assert_frame_equal(
            without["Monthly Summary"], with_exc["Monthly Summary"]
        )


class TestEmptyFrames:
    def test_no_sap_rows_does_not_crash(self, scenario):
        """pandas 3.0: empty .map/.apply blow up. Guard the exceptions path too."""
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap.iloc[0:0], portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        assert sheets["Supplier Summary"].empty
        assert sheets["Should Have Uploaded"].empty


class TestShouldHaveUploaded:
    def test_lists_only_non_exception_zero_upload_suppliers(self, scenario):
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        names = set(sheets["Should Have Uploaded"]["Vendor Name"])

        assert "ACQUA MINERALE SAN BENEDETTO" in names  # expected, uploaded nothing
        assert "AMERICOLD TACOMA" in names              # not on tracker, still chased
        assert "BOTHWELL CHEESE" not in names           # an exception
        assert "GOOD SUPPLIER" not in names             # it uploaded

    def test_sorted_by_most_pos_first(self, scenario):
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        counts = sheets["Should Have Uploaded"]["Inbound POs Expected"].tolist()
        assert counts == sorted(counts, reverse=True)
        # Acqua has 2 POs, Americold 1.
        top = sheets["Should Have Uploaded"].iloc[0]
        assert top["Vendor Name"] == "ACQUA MINERALE SAN BENEDETTO"
        assert top["Inbound POs Expected"] == 2
        assert top["Portal Uploads"] == 0

    def test_no_billback_column(self, scenario):
        """The chase-list is a chase-list, not an invoice. The BB- tabs are the
        billing artefact; duplicating a dollar figure here invites the two
        drifting apart."""
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        assert "Bill-Back Total" not in sheets["Should Have Uploaded"].columns


class TestExemptButSubmitting:
    """Exempt suppliers who upload anyway — their exemption is probably stale."""

    def test_lists_an_exempt_supplier_that_uploaded(self):
        sap = pd.DataFrame([
            _sap_row("4001", "70001111", "BOTHWELL CHEESE"),
            _sap_row("4002", "70001111", "BOTHWELL CHEESE"),
        ])
        portal = pd.DataFrame([_portal_row("4001", "BOTHWELL CHEESE")])
        exceptions = {normalize_supplier_name("BOTHWELL CHEESE"): _rec("BOTHWELL CHEESE")}

        sheet = build_report(sap, portal, 2026, 6, exceptions=exceptions)[
            "Exempt But Submitting"
        ]
        assert list(sheet["Vendor Name"]) == ["BOTHWELL CHEESE"]
        row = sheet.iloc[0]
        assert row["Inbound POs"] == 2
        assert row["Portal Files Uploaded"] == 1
        assert row["Of Which Rejected"] == 0
        assert row["POs Still Missing A File"] == 1

    def test_a_silent_exempt_supplier_is_excluded(self):
        """Uploading nothing is what an exemption is FOR — not a signal."""
        sap = pd.DataFrame([_sap_row("4003", "70001111", "BOTHWELL CHEESE")])
        exceptions = {normalize_supplier_name("BOTHWELL CHEESE"): _rec("BOTHWELL CHEESE")}
        portal = pd.DataFrame(columns=list(_portal_row("x", "y")))

        sheet = build_report(sap, portal, 2026, 6, exceptions=exceptions)[
            "Exempt But Submitting"
        ]
        assert sheet.empty

    def test_a_non_exempt_supplier_is_excluded(self, scenario):
        """GOOD SUPPLIER uploads, but it was never exempt — not this sheet's business."""
        sap, portal, exceptions, tracker = scenario
        sheet = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )["Exempt But Submitting"]
        assert "GOOD SUPPLIER" not in set(sheet["Vendor Name"])

    def test_a_rejected_upload_still_counts_as_submitting(self):
        """An Invalid upload means they engaged with the process — that's the signal."""
        sap = pd.DataFrame([_sap_row("4004", "70001111", "BOTHWELL CHEESE")])
        portal = pd.DataFrame([_portal_row("4004", "BOTHWELL CHEESE", status="Invalid")])
        exceptions = {normalize_supplier_name("BOTHWELL CHEESE"): _rec("BOTHWELL CHEESE")}

        sheet = build_report(sap, portal, 2026, 6, exceptions=exceptions)[
            "Exempt But Submitting"
        ]
        assert len(sheet) == 1
        assert sheet.iloc[0]["Of Which Rejected"] == 1

    def test_empty_sap_does_not_crash(self, scenario):
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap.iloc[0:0], portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        assert sheets["Exempt But Submitting"].empty

    def test_a_supplier_with_an_invalid_upload_is_excluded(self):
        """An Invalid upload still means the supplier knows the process exists.
        They belong in bill-back, not on the 'never even tried' list."""
        sap = pd.DataFrame([_sap_row("2001", "70003333", "TRIED AND FAILED")])
        portal = pd.DataFrame([_portal_row("2001", "TRIED AND FAILED", status="Invalid")])
        sheets = build_report(sap, portal, 2026, 6)
        assert "TRIED AND FAILED" not in set(sheets["Should Have Uploaded"]["Vendor Name"])

    def test_a_supplier_with_no_inbound_delivery_is_excluded(self):
        """No inbound delivery means there was nothing to document yet."""
        sap = pd.DataFrame([_sap_row("3001", "70004444", "NOT SHIPPED YET", inbound="")])
        sheets = build_report(sap, pd.DataFrame(columns=list(_portal_row("x", "y"))), 2026, 6)
        assert "NOT SHIPPED YET" not in set(sheets["Should Have Uploaded"]["Vendor Name"])


from streamlit.testing.v1 import AppTest


class TestDashboardBoots:
    def test_page_1_still_renders(self):
        """The exceptions panel must not crash the page when no DB is configured."""
        at = AppTest.from_file("app.py", default_timeout=30).run()
        assert not at.exception


from src import supplier_exceptions_ui


class TestLoadExceptionsFailsOpen:
    """A dropped Neon connection must never surface the DSN/host, and must
    never break the report -- load_exceptions_or_empty() always returns."""

    def test_fails_open_and_does_not_leak_the_host(self, monkeypatch):
        fake_host = "secret-host.neon.tech"

        class FakeOperationalError(Exception):
            def __str__(self):
                return f"failed to resolve host '{fake_host}': timeout"

        monkeypatch.setattr(
            supplier_exceptions_ui, "_dsn", lambda: "postgresql://user:pw@" + fake_host + "/db"
        )

        def _raise_ensure_schema(self):
            raise FakeOperationalError()

        monkeypatch.setattr(
            supplier_exceptions_ui.ExceptionStore, "ensure_schema", _raise_ensure_schema
        )

        records, tracker_names, message = supplier_exceptions_ui.load_exceptions_or_empty()

        # Fails open: report generation can proceed with an empty exceptions set.
        assert records == {}
        assert tracker_names == set()

        # Never leaks the host (or any part of the DSN) into the message.
        assert message is not None
        assert fake_host not in message
        assert "user:pw" not in message
        assert "FakeOperationalError" in message
