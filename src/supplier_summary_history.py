"""Supplier Summary history — data + persistence layer (no Streamlit).

The Supplier Summary sheet (``compliance_engine._supplier_summary``) is one
run's snapshot: whatever SAP + Portal files were uploaded, scoped to one
report Year/Month. This module lets the Supplier Summary Dashboard **save**
a run under a chosen snapshot date and pull that history back out again, so
a weekly check-in ("how did this vendor's compliance move since last week?")
doesn't require re-uploading and recomputing every prior week's files.

Mirrors ``column_variants``'s store shape (short-lived autocommit Postgres
connections, ``ensure_schema()``, graceful absence when the DB isn't
configured) — same hosted Postgres, one more table.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime

import pandas as pd
import psycopg
from psycopg.rows import dict_row

# ---------------------------------------------------------------------------
# Pure transform: Supplier Summary DataFrame -> snapshot rows
# ---------------------------------------------------------------------------
SS_VENDOR_NUMBER   = "Vendor Number"
SS_VENDOR_NAME     = "Vendor Name"
SS_BDM             = "BDM"
SS_BDM_DESCRIPTION = "BDM Description"
SS_EXCEPTION       = "Exception Status"
SS_TOTAL_POS       = "Total SAP POs"
SS_WITH_INBOUND    = "SAP POs With Inbound Delivery"
SS_FOUND           = "Portal Files Found"
SS_MISSING         = "Missing Portal Files"
SS_INVALID         = "Invalid Portal Uploads"
SS_NO_SAP_INBOUND  = "Portal Files With No SAP Inbound"
SS_CLOSED          = "Closed POs"
SS_PROCESSING      = "Processing POs"

REQUIRED_SUPPLIER_SUMMARY_COLS = [
    SS_VENDOR_NUMBER, SS_VENDOR_NAME, SS_WITH_INBOUND, SS_FOUND,
    SS_MISSING, SS_INVALID,
]


def summary_to_snapshot_rows(
    supplier_summary: pd.DataFrame,
    snapshot_date: date | datetime,
    report_year: int,
    report_month: int,
) -> list[dict]:
    """Convert one Supplier Summary DataFrame into snapshot rows ready to
    save. Compliance percentage is recomputed from the numeric found/with-
    inbound counts rather than parsed from the sheet's formatted "82.3%"
    string, so it stays an exact float. A vendor with no inbound POs that
    month gets ``compliance_pct = None`` (nothing to be compliant about),
    matching the sheet's own 0.0 display-only convention but staying
    distinguishable in stored history."""
    missing = [c for c in REQUIRED_SUPPLIER_SUMMARY_COLS if c not in supplier_summary.columns]
    if missing:
        raise ValueError(f"Supplier Summary is missing column(s): {', '.join(missing)}")

    snap = snapshot_date.date() if isinstance(snapshot_date, datetime) else snapshot_date

    rows = []
    for _, r in supplier_summary.iterrows():
        with_inbound = int(r[SS_WITH_INBOUND])
        found = int(r[SS_FOUND])
        pct = (found / with_inbound) if with_inbound else None
        rows.append({
            "snapshot_date": snap,
            "report_year": int(report_year),
            "report_month": int(report_month),
            "vendor_number": str(r[SS_VENDOR_NUMBER]),
            "vendor_name": str(r[SS_VENDOR_NAME]),
            "bdm": str(r.get(SS_BDM, "") or ""),
            "bdm_description": str(r.get(SS_BDM_DESCRIPTION, "") or ""),
            "exception_status": str(r.get(SS_EXCEPTION, "") or ""),
            "total_sap_pos": int(r[SS_TOTAL_POS]),
            "sap_pos_with_inbound": with_inbound,
            "portal_files_found": found,
            "missing_portal_files": int(r[SS_MISSING]),
            "invalid_portal_uploads": int(r[SS_INVALID]),
            "portal_files_no_sap_inbound": int(r.get(SS_NO_SAP_INBOUND, 0) or 0),
            "closed_pos": int(r.get(SS_CLOSED, 0) or 0),
            "processing_pos": int(r.get(SS_PROCESSING, 0) or 0),
            "compliance_pct": pct,
        })
    return rows


# ---------------------------------------------------------------------------
# Model
# ---------------------------------------------------------------------------
@dataclass
class SnapshotRow:
    snapshot_date: date
    report_year: int
    report_month: int
    vendor_number: str
    vendor_name: str
    bdm: str
    bdm_description: str
    exception_status: str
    total_sap_pos: int
    sap_pos_with_inbound: int
    portal_files_found: int
    missing_portal_files: int
    invalid_portal_uploads: int
    portal_files_no_sap_inbound: int
    closed_pos: int
    processing_pos: int
    compliance_pct: float | None


# ---------------------------------------------------------------------------
# Store
# ---------------------------------------------------------------------------
_SCHEMA_STATEMENTS = (
    """
    CREATE TABLE IF NOT EXISTS supplier_summary_snapshots (
        id                            BIGSERIAL   PRIMARY KEY,
        snapshot_date                 DATE        NOT NULL,
        report_year                   INT         NOT NULL,
        report_month                  INT         NOT NULL,
        vendor_number                 TEXT        NOT NULL,
        vendor_name                   TEXT        NOT NULL,
        bdm                           TEXT        NOT NULL DEFAULT '',
        bdm_description               TEXT        NOT NULL DEFAULT '',
        exception_status               TEXT        NOT NULL DEFAULT '',
        total_sap_pos                 INT         NOT NULL DEFAULT 0,
        sap_pos_with_inbound           INT         NOT NULL DEFAULT 0,
        portal_files_found            INT         NOT NULL DEFAULT 0,
        missing_portal_files          INT         NOT NULL DEFAULT 0,
        invalid_portal_uploads        INT         NOT NULL DEFAULT 0,
        portal_files_no_sap_inbound   INT         NOT NULL DEFAULT 0,
        closed_pos                    INT         NOT NULL DEFAULT 0,
        processing_pos                INT         NOT NULL DEFAULT 0,
        compliance_pct                DOUBLE PRECISION,
        created_at                    TIMESTAMPTZ NOT NULL DEFAULT now()
    )
    """,
    "CREATE UNIQUE INDEX IF NOT EXISTS ux_supplier_summary_snapshots_date_vendor "
    "ON supplier_summary_snapshots (snapshot_date, vendor_number)",
    "CREATE INDEX IF NOT EXISTS ix_supplier_summary_snapshots_vendor "
    "ON supplier_summary_snapshots (vendor_number, snapshot_date)",
)

_COLS = (
    "snapshot_date, report_year, report_month, vendor_number, vendor_name, "
    "bdm, bdm_description, exception_status, total_sap_pos, "
    "sap_pos_with_inbound, portal_files_found, missing_portal_files, "
    "invalid_portal_uploads, portal_files_no_sap_inbound, closed_pos, "
    "processing_pos, compliance_pct"
)


def _row_to_snapshot(row: dict) -> SnapshotRow:
    return SnapshotRow(**{k: row[k] for k in SnapshotRow.__dataclass_fields__})


class SupplierSummaryHistoryStore:
    """Postgres-backed store for weekly Supplier Summary snapshots.

    One short-lived autocommit connection per operation — same reasoning as
    ``column_variants.VariantStore``: robust against serverless Postgres
    (Neon) closing idle connections."""

    def __init__(self, dsn: str):
        self.dsn = dsn

    def _connect(self):
        return psycopg.connect(self.dsn, autocommit=True, row_factory=dict_row)

    def ensure_schema(self) -> None:
        with self._connect() as conn:
            for stmt in _SCHEMA_STATEMENTS:
                conn.execute(stmt)

    def save_snapshot(self, rows: list[dict]) -> int:
        """Upsert snapshot rows (one per vendor). Re-saving the same
        snapshot_date + vendor overwrites — a re-run on the same day replaces
        rather than duplicates. Returns the number of rows written."""
        if not rows:
            return 0
        cols = list(rows[0].keys())
        placeholders = ", ".join(f"%({c})s" for c in cols)
        col_list = ", ".join(cols)
        update_list = ", ".join(
            f"{c} = EXCLUDED.{c}" for c in cols
            if c not in ("snapshot_date", "vendor_number")
        )
        sql = (
            f"INSERT INTO supplier_summary_snapshots ({col_list}) "
            f"VALUES ({placeholders}) "
            f"ON CONFLICT (snapshot_date, vendor_number) DO UPDATE SET {update_list}"
        )
        with self._connect() as conn, conn.cursor() as cur:
            cur.executemany(sql, rows)
        return len(rows)

    def list_snapshot_dates(self) -> list[date]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT snapshot_date FROM supplier_summary_snapshots "
                "ORDER BY snapshot_date"
            ).fetchall()
        return [r["snapshot_date"] for r in rows]

    def history(
        self,
        vendor_numbers: list[str] | None = None,
        since: date | None = None,
    ) -> list[SnapshotRow]:
        """All snapshot rows, optionally filtered to specific vendors and/or
        a start date. Ordered by vendor then date for easy trend-building."""
        clauses, params = [], []
        if vendor_numbers:
            clauses.append("vendor_number = ANY(%s)")
            params.append(list(vendor_numbers))
        if since is not None:
            clauses.append("snapshot_date >= %s")
            params.append(since)
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT {_COLS} FROM supplier_summary_snapshots {where} "
                f"ORDER BY vendor_number, snapshot_date",
                params,
            ).fetchall()
        return [_row_to_snapshot(r) for r in rows]

    def latest_two_dates(self) -> tuple[date, date] | None:
        """The two most recent distinct snapshot dates, newest first — the
        pair the week-over-week movers table compares. None if fewer than
        two snapshots have ever been saved."""
        dates = self.list_snapshot_dates()
        if len(dates) < 2:
            return None
        return dates[-1], dates[-2]
