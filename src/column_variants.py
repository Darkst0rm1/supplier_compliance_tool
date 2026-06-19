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
import psycopg
from psycopg.errors import UniqueViolation
from psycopg.rows import dict_row
from psycopg.types.json import Jsonb

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


# --- store -----------------------------------------------------------------
_SCHEMA_STATEMENTS = (
    """
    CREATE TABLE IF NOT EXISTS column_variants (
        id          BIGSERIAL   PRIMARY KEY,
        report_key  TEXT        NOT NULL CHECK (
                        report_key IN ('delivery_shortage', 'sales_order_unconfirmed')
                    ),
        name        TEXT        NOT NULL,
        columns     JSONB       NOT NULL CHECK (
                        jsonb_typeof(columns) = 'array'
                        AND jsonb_array_length(columns) > 0
                    ),
        created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
        updated_at  TIMESTAMPTZ NOT NULL DEFAULT now()
    )
    """,
    "CREATE INDEX IF NOT EXISTS ix_column_variants_report ON column_variants (report_key)",
    "CREATE UNIQUE INDEX IF NOT EXISTS ux_column_variants_report_name_ci "
    "ON column_variants (report_key, LOWER(name))",
)

_COLS = "id, report_key, name, columns, created_at, updated_at"


def _row_to_variant(row: dict) -> Variant:
    return Variant(
        id=row["id"],
        report_key=row["report_key"],
        name=row["name"],
        columns=list(row["columns"]),
        created_at=row.get("created_at"),
        updated_at=row.get("updated_at"),
    )


class VariantStore:
    """Postgres-backed store for shared column variants.

    One short-lived autocommit connection is opened per operation, which is
    robust against serverless Postgres (Neon) closing idle connections.
    """

    def __init__(self, dsn: str):
        self.dsn = dsn

    def _connect(self):
        return psycopg.connect(self.dsn, autocommit=True, row_factory=dict_row)

    def ensure_schema(self) -> None:
        with self._connect() as conn:
            for stmt in _SCHEMA_STATEMENTS:
                conn.execute(stmt)

    def list_variants(self, report_key: str) -> list[Variant]:
        validate_report_key(report_key)
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT {_COLS} FROM column_variants "
                f"WHERE report_key = %s ORDER BY LOWER(name)",
                (report_key,),
            ).fetchall()
        return [_row_to_variant(r) for r in rows]

    def get_variant(self, variant_id: int) -> Variant:
        with self._connect() as conn:
            row = conn.execute(
                f"SELECT {_COLS} FROM column_variants WHERE id = %s", (variant_id,)
            ).fetchone()
        if row is None:
            raise VariantNotFoundError(f"variant {variant_id} not found")
        return _row_to_variant(row)

    def create_variant(self, report_key: str, name: str, columns: Any) -> Variant:
        validate_report_key(report_key)
        clean_name = validate_name(name)
        clean_cols = normalize_columns(columns)
        try:
            with self._connect() as conn:
                row = conn.execute(
                    f"INSERT INTO column_variants (report_key, name, columns) "
                    f"VALUES (%s, %s, %s) RETURNING {_COLS}",
                    (report_key, clean_name, Jsonb(clean_cols)),
                ).fetchone()
        except UniqueViolation as exc:
            raise DuplicateVariantError(
                f"a variant named '{clean_name}' already exists"
            ) from exc
        return _row_to_variant(row)

    def update_columns(self, variant_id: int, columns: Any) -> Variant:
        clean_cols = normalize_columns(columns)
        with self._connect() as conn:
            row = conn.execute(
                f"UPDATE column_variants SET columns = %s, updated_at = now() "
                f"WHERE id = %s RETURNING {_COLS}",
                (Jsonb(clean_cols), variant_id),
            ).fetchone()
        if row is None:
            raise VariantNotFoundError(f"variant {variant_id} not found")
        return _row_to_variant(row)

    def rename_variant(self, variant_id: int, name: str) -> Variant:
        clean_name = validate_name(name)
        try:
            with self._connect() as conn:
                row = conn.execute(
                    f"UPDATE column_variants SET name = %s, updated_at = now() "
                    f"WHERE id = %s RETURNING {_COLS}",
                    (clean_name, variant_id),
                ).fetchone()
        except UniqueViolation as exc:
            raise DuplicateVariantError(
                f"a variant named '{clean_name}' already exists"
            ) from exc
        if row is None:
            raise VariantNotFoundError(f"variant {variant_id} not found")
        return _row_to_variant(row)

    def delete_variant(self, variant_id: int) -> None:
        with self._connect() as conn:
            cur = conn.execute("DELETE FROM column_variants WHERE id = %s", (variant_id,))
        if cur.rowcount == 0:
            raise VariantNotFoundError(f"variant {variant_id} not found")
