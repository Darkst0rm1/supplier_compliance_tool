"""Tests for shared column variants."""
from __future__ import annotations

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
