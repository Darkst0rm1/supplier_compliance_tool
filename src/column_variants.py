"""Shared Column Variants — data + persistence layer (no Streamlit).

A *variant* is a named, ordered subset of columns for one report's detail table.
"Standard" (all columns) is built in and never stored. User variants are shared
(no per-user scoping) and live in a hosted Postgres table, keyed by report_key.
"""
from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from datetime import datetime
from typing import Any

import pandas as pd

# --- constants -------------------------------------------------------------
REPORT_DELIVERY_SHORTAGE = "delivery_shortage"
REPORT_SALES_ORDER_UNCONFIRMED = "sales_order_unconfirmed"
VALID_REPORT_KEYS = (REPORT_DELIVERY_SHORTAGE, REPORT_SALES_ORDER_UNCONFIRMED)

STANDARD_NAME = "Standard"
MAX_NAME_LEN = 120


# --- errors ----------------------------------------------------------------
class VariantError(Exception):
    """Base error for the column-variants feature."""


class VariantValidationError(VariantError):
    """Invalid name, report_key, or column list."""


class VariantNotFoundError(VariantError):
    """No variant with the given id."""


class DuplicateVariantError(VariantError):
    """A variant with the same name (case-insensitive) already exists."""


# --- model -----------------------------------------------------------------
@dataclass
class Variant:
    id: int
    report_key: str
    name: str
    columns: list[str]
    created_at: datetime | None = None
    updated_at: datetime | None = None


# --- pure validation / helpers --------------------------------------------
def is_reserved_name(name: Any) -> bool:
    return str(name).strip().casefold() == STANDARD_NAME.casefold()


def validate_name(name: Any) -> str:
    if not isinstance(name, str):
        raise VariantValidationError("name must be a string")
    cleaned = name.strip()
    if not cleaned:
        raise VariantValidationError("name cannot be empty")
    if len(cleaned) > MAX_NAME_LEN:
        raise VariantValidationError(f"name cannot exceed {MAX_NAME_LEN} characters")
    if is_reserved_name(cleaned):
        raise VariantValidationError(f"'{STANDARD_NAME}' is a reserved name")
    return cleaned


def validate_report_key(report_key: Any) -> str:
    if report_key not in VALID_REPORT_KEYS:
        raise VariantValidationError(
            f"report_key must be one of {VALID_REPORT_KEYS}, got {report_key!r}"
        )
    return report_key


def normalize_columns(columns: Any) -> list[str]:
    """Coerce to a de-duplicated, order-preserving list of non-empty strings."""
    if isinstance(columns, str) or not isinstance(columns, Iterable):
        raise VariantValidationError("columns must be a list")
    out: list[str] = []
    seen: set[str] = set()
    for c in columns:
        s = str(c).strip()
        if not s or s in seen:
            continue
        seen.add(s)
        out.append(s)
    if not out:
        raise VariantValidationError("a variant must have at least one column")
    return out


def apply_columns(df: pd.DataFrame, columns: Iterable[str] | None) -> pd.DataFrame:
    """Project & order ``df`` to ``columns``. Missing columns are skipped; if
    none of the requested columns are present (or ``columns`` is falsy), ``df``
    is returned unchanged (the "Standard" behaviour)."""
    if not columns:
        return df
    present = [c for c in columns if c in df.columns]
    if not present:
        return df
    return df[present]
