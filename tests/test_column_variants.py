"""Tests for shared column variants."""
from __future__ import annotations

import os
import time

import pandas as pd
import pytest

from src.column_variants import (
    MAX_NAME_LEN,
    REPORT_DELIVERY_SHORTAGE,
    REPORT_SALES_ORDER_UNCONFIRMED,
    STANDARD_NAME,
    VALID_REPORT_KEYS,
    VariantValidationError,
    apply_columns,
    is_reserved_name,
    normalize_columns,
    validate_name,
    validate_report_key,
)


# -- validate_name -----------------------------------------------------------
def test_validate_name_trims_and_returns():
    assert validate_name("  My View  ") == "My View"


def test_validate_name_rejects_empty():
    with pytest.raises(VariantValidationError):
        validate_name("   ")


def test_validate_name_rejects_too_long():
    with pytest.raises(VariantValidationError):
        validate_name("x" * (MAX_NAME_LEN + 1))


def test_validate_name_rejects_reserved_any_case():
    with pytest.raises(VariantValidationError):
        validate_name("standard")
    with pytest.raises(VariantValidationError):
        validate_name("  STANDARD ")


# -- is_reserved_name --------------------------------------------------------
def test_is_reserved_name_case_insensitive():
    assert is_reserved_name("Standard")
    assert is_reserved_name(" sTaNdArD ")
    assert not is_reserved_name("My View")


# -- validate_report_key -----------------------------------------------------
def test_validate_report_key_accepts_known():
    assert validate_report_key(REPORT_DELIVERY_SHORTAGE) == REPORT_DELIVERY_SHORTAGE
    assert REPORT_SALES_ORDER_UNCONFIRMED in VALID_REPORT_KEYS


def test_validate_report_key_rejects_unknown():
    with pytest.raises(VariantValidationError):
        validate_report_key("nope")


# -- normalize_columns -------------------------------------------------------
def test_normalize_columns_dedupes_preserving_order():
    assert normalize_columns(["a", "b", "a", " c "]) == ["a", "b", "c"]


def test_normalize_columns_drops_blanks():
    assert normalize_columns(["a", "", "  ", "b"]) == ["a", "b"]


def test_normalize_columns_rejects_empty_result():
    with pytest.raises(VariantValidationError):
        normalize_columns([])
    with pytest.raises(VariantValidationError):
        normalize_columns(["", "   "])


def test_normalize_columns_rejects_non_list():
    with pytest.raises(VariantValidationError):
        normalize_columns("abc")


# -- apply_columns -----------------------------------------------------------
def _df():
    return pd.DataFrame({"a": [1], "b": [2], "c": [3]})


def test_apply_columns_projects_and_orders():
    out = apply_columns(_df(), ["c", "a"])
    assert list(out.columns) == ["c", "a"]


def test_apply_columns_skips_missing():
    out = apply_columns(_df(), ["c", "zzz", "a"])
    assert list(out.columns) == ["c", "a"]


def test_apply_columns_none_present_returns_unchanged():
    out = apply_columns(_df(), ["x", "y"])
    assert list(out.columns) == ["a", "b", "c"]


def test_apply_columns_empty_or_none_returns_unchanged():
    assert list(apply_columns(_df(), []).columns) == ["a", "b", "c"]
    assert list(apply_columns(_df(), None).columns) == ["a", "b", "c"]


def test_standard_name_constant():
    assert STANDARD_NAME == "Standard"


from src.column_variants import (
    DuplicateVariantError,
    VariantNotFoundError,
    VariantStore,
)

TEST_DSN = os.environ.get("TEST_DATABASE_URL")
needs_db = pytest.mark.skipif(not TEST_DSN, reason="TEST_DATABASE_URL not set")
PREFIX = "pytest_cv_"


@pytest.fixture
def store():
    s = VariantStore(TEST_DSN)
    s.ensure_schema()
    yield s
    with s._connect() as conn:  # noqa: SLF001 - test cleanup
        conn.execute("DELETE FROM column_variants WHERE name LIKE %s", (PREFIX + "%",))


@needs_db
def test_create_list_get_roundtrip(store):
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "A", ["product", "plant"])
    assert v.id > 0
    assert v.columns == ["product", "plant"]
    fetched = store.get_variant(v.id)
    assert fetched.name == PREFIX + "A"
    names = [x.name for x in store.list_variants(REPORT_DELIVERY_SHORTAGE)]
    assert PREFIX + "A" in names


@needs_db
def test_duplicate_name_case_insensitive(store):
    store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Dup", ["product"])
    with pytest.raises(DuplicateVariantError):
        store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "dup", ["plant"])


@needs_db
def test_update_columns_bumps_updated_at(store):
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Upd", ["product"])
    time.sleep(0.02)
    v2 = store.update_columns(v.id, ["plant", "short_amount"])
    assert v2.columns == ["plant", "short_amount"]
    assert v2.updated_at > v2.created_at


@needs_db
def test_rename(store):
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Old", ["product"])
    v2 = store.rename_variant(v.id, PREFIX + "New")
    assert v2.name == PREFIX + "New"


@needs_db
def test_rename_into_existing_name_rejected(store):
    store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Taken", ["product"])
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Move", ["plant"])
    with pytest.raises(DuplicateVariantError):
        store.rename_variant(v.id, PREFIX + "taken")


@needs_db
def test_delete(store):
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Gone", ["product"])
    store.delete_variant(v.id)
    with pytest.raises(VariantNotFoundError):
        store.get_variant(v.id)


@needs_db
def test_variants_separated_by_report_key(store):
    store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "DelOnly", ["product"])
    so_names = [x.name for x in store.list_variants(REPORT_SALES_ORDER_UNCONFIRMED)]
    assert PREFIX + "DelOnly" not in so_names


@needs_db
def test_delete_missing_id_raises(store):
    with pytest.raises(VariantNotFoundError):
        store.delete_variant(999999999)
