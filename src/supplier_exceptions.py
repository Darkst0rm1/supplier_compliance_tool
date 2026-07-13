"""Supplier exceptions -- data + persistence layer (no Streamlit).

A supplier with an approved *exception* is not required to upload inbound
documentation. The list is seeded from the Master Inbound Delivery Compliance
Tracker, then owned by this table.

Mirrors src/column_variants.py: psycopg3, one short-lived autocommit connection
per operation (robust against Neon closing idle connections), shared with no
per-user scoping.

IMPORTANT: this feature is informational. It does NOT change bill-back or the
compliance percentage, so a DB outage cannot wrongly excuse a supplier from a
charge -- it only removes an annotation.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any

from .config import (
    EXCEPTION_STATUS_EXCEPTION,
    EXCEPTION_STATUS_EXPECTED,
    EXCEPTION_STATUS_NOT_ON_TRACKER,
    REASON_EXEMPT_MARK,
    REASON_MANUAL,
    REASON_UNABLE_TO_COMPLY,
)
from .normalizer import normalize_supplier_name

MAX_NAME_LEN = 200
VALID_REASONS = (REASON_UNABLE_TO_COMPLY, REASON_EXEMPT_MARK, REASON_MANUAL)


# --- errors ----------------------------------------------------------------
class ExceptionError(Exception):
    """Base error for the supplier-exceptions feature."""


class ExceptionValidationError(ExceptionError):
    """Invalid supplier name or reason."""


class ExceptionNotFoundError(ExceptionError):
    """No exception with the given normalized name."""


class DuplicateExceptionError(ExceptionError):
    """This supplier is already an exception."""


# --- model -----------------------------------------------------------------
@dataclass
class ExceptionRecord:
    id: int
    supplier_name: str
    normalized_name: str
    vendor_number: str | None
    reason: str
    added_by: str | None = None
    added_at: datetime | None = None


# --- pure validation / classification --------------------------------------
def validate_supplier_name(name: Any) -> str:
    if not isinstance(name, str):
        raise ExceptionValidationError("supplier name must be a string")
    cleaned = name.strip()
    if not cleaned:
        raise ExceptionValidationError("supplier name cannot be empty")
    if len(cleaned) > MAX_NAME_LEN:
        raise ExceptionValidationError(
            f"supplier name cannot exceed {MAX_NAME_LEN} characters"
        )
    return cleaned


def validate_reason(reason: Any) -> str:
    if reason not in VALID_REASONS:
        raise ExceptionValidationError(
            f"reason must be one of {VALID_REASONS}, got {reason!r}"
        )
    return reason


def classify_supplier(
    vendor_name: Any,
    vendor_number: Any,
    exceptions: dict[str, ExceptionRecord],
    tracker_names: set[str],
) -> str:
    """Return the Exception Status for one SAP vendor.

    Matches on vendor number first (exact, when we have one on file), then on the
    normalized name -- the tracker's own Supplier # column is useless as a key
    (G61 / 491 vs SAP's 8-digit 70007212, zero overlap), so name is usually all
    there is.

    `tracker_names` is every supplier the tracker knows about, exception or not.
    Pass an empty set to collapse "Not on tracker" into "Expected to upload".
    """
    number = "" if vendor_number is None else str(vendor_number).strip()
    if number:
        for record in exceptions.values():
            if record.vendor_number and record.vendor_number.strip() == number:
                return EXCEPTION_STATUS_EXCEPTION

    key = normalize_supplier_name(vendor_name)
    if key and key in exceptions:
        return EXCEPTION_STATUS_EXCEPTION
    if not tracker_names:
        return EXCEPTION_STATUS_EXPECTED
    if key and key in tracker_names:
        return EXCEPTION_STATUS_EXPECTED
    return EXCEPTION_STATUS_NOT_ON_TRACKER


# --- store -----------------------------------------------------------------
_SCHEMA_STATEMENTS = (
    """
    CREATE TABLE IF NOT EXISTS supplier_exceptions (
        id              BIGSERIAL   PRIMARY KEY,
        supplier_name   TEXT        NOT NULL CHECK (length(btrim(supplier_name)) > 0),
        normalized_name TEXT        NOT NULL CHECK (length(btrim(normalized_name)) > 0),
        vendor_number   TEXT,
        reason          TEXT        NOT NULL CHECK (
                            reason IN ('Unable to Comply', 'EXEMPT mark', 'Manual')
                        ),
        added_by        TEXT,
        added_at        TIMESTAMPTZ NOT NULL DEFAULT now()
    )
    """,
    "CREATE UNIQUE INDEX IF NOT EXISTS ux_supplier_exceptions_normalized "
    "ON supplier_exceptions (normalized_name)",
)

_COLS = "id, supplier_name, normalized_name, vendor_number, reason, added_by, added_at"


def _row_to_record(row: dict) -> ExceptionRecord:
    """Build a record, recomputing normalized_name from the display name.

    The stored `normalized_name` column can go stale if `normalize_supplier_name`
    changes after the row was written (it already has once, for diacritics). The
    column is still the DB's unique index and is used to target `remove_exception`,
    but it must never be trusted as the match key -- `classify_supplier` always
    recomputes it from `supplier_name`, so the record must agree.
    """
    return ExceptionRecord(
        id=row["id"],
        supplier_name=row["supplier_name"],
        normalized_name=normalize_supplier_name(row["supplier_name"]),
        vendor_number=row.get("vendor_number"),
        reason=row["reason"],
        added_by=row.get("added_by"),
        added_at=row.get("added_at"),
    )


class ExceptionStore:
    """Postgres-backed store for the shared supplier exceptions list."""

    def __init__(self, dsn: str):
        self.dsn = dsn

    def _connect(self):
        # Imported lazily so that importing this module for its pure functions
        # (classify_supplier, ExceptionRecord, validators) never requires
        # psycopg -- compliance_engine imports this module and must never break
        # at import time on a psycopg install/binary problem.
        import psycopg
        from psycopg.rows import dict_row

        return psycopg.connect(self.dsn, autocommit=True, row_factory=dict_row)

    def ensure_schema(self) -> None:
        with self._connect() as conn:
            for stmt in _SCHEMA_STATEMENTS:
                conn.execute(stmt)

    def load_exceptions(self) -> dict[str, ExceptionRecord]:
        """All exceptions, keyed by normalize_supplier_name(supplier_name).

        Recomputed rather than trusted from the stored `normalized_name` column,
        so the dict always matches what `classify_supplier` looks up -- even if
        the normalizer changed after a row was written.
        """
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT {_COLS} FROM supplier_exceptions ORDER BY supplier_name"
            ).fetchall()
        records = [_row_to_record(r) for r in rows]
        return {rec.normalized_name: rec for rec in records}

    def add_exception(
        self,
        supplier_name: str,
        reason: str = REASON_MANUAL,
        vendor_number: str | None = None,
        added_by: str | None = None,
    ) -> ExceptionRecord:
        clean_name = validate_supplier_name(supplier_name)
        clean_reason = validate_reason(reason)
        key = normalize_supplier_name(clean_name)
        if not key:
            raise ExceptionValidationError(
                f"'{supplier_name}' normalizes to an empty key"
            )
        from psycopg.errors import UniqueViolation

        try:
            with self._connect() as conn:
                row = conn.execute(
                    "INSERT INTO supplier_exceptions "
                    "(supplier_name, normalized_name, vendor_number, reason, added_by) "
                    f"VALUES (%s, %s, %s, %s, %s) RETURNING {_COLS}",
                    (clean_name, key, vendor_number, clean_reason, added_by),
                ).fetchone()
        except UniqueViolation as exc:
            raise DuplicateExceptionError(
                f"'{clean_name}' is already an exception"
            ) from exc
        return _row_to_record(row)

    def remove_exception(self, normalized_name: str) -> None:
        with self._connect() as conn:
            cur = conn.execute(
                "DELETE FROM supplier_exceptions WHERE normalized_name = %s",
                (normalized_name,),
            )
        if cur.rowcount == 0:
            raise ExceptionNotFoundError(f"'{normalized_name}' is not an exception")
