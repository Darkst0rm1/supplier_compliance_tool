"""Streamlit glue for the supplier exceptions list.

Kept separate from src/supplier_exceptions.py so the store stays importable
(and testable) without Streamlit -- the same split column_variants uses.
"""
from __future__ import annotations

import os

import streamlit as st

from .supplier_exceptions import (
    DuplicateExceptionError,
    ExceptionNotFoundError,
    ExceptionRecord,
    ExceptionStore,
    ExceptionValidationError,
)
from .config import REASON_MANUAL


def _dsn() -> str | None:
    """Read the Postgres DSN from Streamlit secrets, then the env var. Never hardcode."""
    try:
        return st.secrets["postgres"]["dsn"]
    except Exception:  # noqa: BLE001 - secrets may be absent entirely
        return os.environ.get("DATABASE_URL")


def load_exceptions_or_empty() -> tuple[dict[str, ExceptionRecord], set[str], str | None]:
    """Load the exceptions list, failing OPEN.

    On any failure returns ({}, set(), message). The report still generates; the
    Exception column just reads "Expected to upload" for everyone. Because this
    feature does not affect bill-back or the compliance percentage, an outage
    cannot wrongly excuse a supplier from a charge.
    """
    dsn = _dsn()
    if not dsn:
        return {}, set(), (
            "No database configured, so supplier exceptions are unavailable. "
            "The report is otherwise complete."
        )
    try:
        store = ExceptionStore(dsn)
        store.ensure_schema()
        return store.load_exceptions(), set(), None
    except Exception as exc:  # noqa: BLE001 - never break the report over this
        return {}, set(), f"Could not load supplier exceptions: {exc}"


def render_exception_manager() -> None:
    """A collapsed panel to add/remove exceptions. Bulk editing stays in Excel."""
    dsn = _dsn()
    with st.expander("Manage Supplier Exceptions"):
        if not dsn:
            st.info(
                "No database configured. Set `[postgres] dsn` in Streamlit secrets "
                "to manage the exceptions list."
            )
            return

        store = ExceptionStore(dsn)
        try:
            store.ensure_schema()
            records = store.load_exceptions()
        except Exception as exc:  # noqa: BLE001
            st.warning(f"Could not reach the exceptions database: {exc}")
            return

        st.caption(
            f"{len(records)} supplier(s) are exempt from uploading inbound "
            "documentation. This list is informational — it does **not** change "
            "bill-back or the compliance percentage."
        )

        if records:
            st.dataframe(
                [
                    {
                        "Supplier": r.supplier_name,
                        "Reason": r.reason,
                        "Vendor Number": r.vendor_number or "—",
                    }
                    for r in records.values()
                ],
                use_container_width=True,
                hide_index=True,
            )

        with st.form("add_exception", clear_on_submit=True):
            name = st.text_input("Supplier name")
            vendor_number = st.text_input("Vendor number (optional)")
            if st.form_submit_button("Add exception") and name.strip():
                try:
                    store.add_exception(
                        name, REASON_MANUAL, vendor_number.strip() or None
                    )
                    st.success(f"Added '{name.strip()}'.")
                    st.rerun()
                except (DuplicateExceptionError, ExceptionValidationError) as exc:
                    st.error(str(exc))

        if records:
            to_remove = st.selectbox(
                "Remove an exception",
                options=[""] + sorted(records),
                format_func=lambda k: "—" if not k else records[k].supplier_name,
            )
            if st.button("Remove", disabled=not to_remove):
                try:
                    store.remove_exception(to_remove)
                    st.success("Removed.")
                    st.rerun()
                except ExceptionNotFoundError as exc:
                    st.error(str(exc))
