"""Seed supplier_exceptions from the Master Inbound Delivery Compliance Tracker.

Idempotent: re-running it adds only suppliers that aren't already there, so it
doubles as the "the tracker changed, pull the new ones in" tool. It never
removes or overwrites -- once seeded, the DB is the source of truth and in-app
edits win.

Usage:
    python scripts/seed_supplier_exceptions.py "C:\\path\\to\\Master ... Tracker.xlsx"
    python scripts/seed_supplier_exceptions.py <path> --dry-run
"""
from __future__ import annotations

import argparse
import os
import sys
import tomllib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.supplier_exceptions import (  # noqa: E402
    DuplicateExceptionError,
    ExceptionStore,
)
from src.tracker_importer import read_tracker_exceptions  # noqa: E402


def _dsn() -> str:
    secrets = Path(".streamlit/secrets.toml")
    if secrets.exists():
        with secrets.open("rb") as fh:
            data = tomllib.load(fh)
        dsn = data.get("postgres", {}).get("dsn")
        if dsn:
            return dsn
    dsn = os.environ.get("DATABASE_URL")
    if not dsn:
        sys.exit(
            "No Postgres DSN. Set [postgres] dsn in .streamlit/secrets.toml or "
            "the DATABASE_URL env var."
        )
    return dsn


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("tracker", help="Path to the tracker .xlsx")
    parser.add_argument(
        "--dry-run", action="store_true", help="Print what would be added, change nothing"
    )
    args = parser.parse_args()

    rows = read_tracker_exceptions(args.tracker)
    print(f"Found {len(rows)} exception supplier(s) in the tracker.")

    if args.dry_run:
        for name, reason in sorted(rows):
            print(f"  would add: {name}  [{reason}]")
        return

    store = ExceptionStore(_dsn())
    store.ensure_schema()

    added = skipped = 0
    for name, reason in sorted(rows):
        try:
            store.add_exception(name, reason)
            print(f"  added:   {name}  [{reason}]")
            added += 1
        except DuplicateExceptionError:
            print(f"  exists:  {name}")
            skipped += 1

    print(f"\nDone. {added} added, {skipped} already present.")
    print(f"Table now holds {len(store.load_exceptions())} exception(s).")


if __name__ == "__main__":
    main()
