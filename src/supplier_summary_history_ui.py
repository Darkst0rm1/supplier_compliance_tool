"""Supplier Summary history — Postgres connection plumbing for the dashboard.

Same DSN lookup and graceful-absence pattern as ``column_variants_ui``: read
from Streamlit secrets, then the ``DATABASE_URL`` env var; never hardcode.
Kept separate from the page so ``pages/9_Supplier_Summary_Dashboard.py``
builds its own charts/tables directly (this module is connection plumbing
only, not a shared widget — unlike the Column Variants panel, this history
view is used on one page).
"""
from __future__ import annotations

import os

import streamlit as st

from src.supplier_summary_history import SupplierSummaryHistoryStore


def dsn() -> str | None:
    """Read the Postgres DSN from Streamlit secrets, then the env var."""
    try:
        return st.secrets["postgres"]["dsn"]
    except Exception:  # noqa: BLE001 - secrets may be absent entirely
        return os.environ.get("DATABASE_URL")


@st.cache_resource(show_spinner=False)
def get_store(dsn_value: str) -> SupplierSummaryHistoryStore:
    store = SupplierSummaryHistoryStore(dsn_value)
    store.ensure_schema()
    return store
