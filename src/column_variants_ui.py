"""Shared Column Variants — Streamlit panel.

`render_variant_panel(report_key, all_columns, key_prefix)` draws the Variant
picker above a detail table and returns the ordered list of columns the table
should display (and the Excel export should use). Falls back to "Standard"
(all columns) whenever the database is not configured or unreachable.
"""
from __future__ import annotations

import os

import streamlit as st

from src.column_variants import (
    STANDARD_NAME,
    VariantError,
    VariantStore,
)


def _dsn() -> str | None:
    """Read the Postgres DSN from Streamlit secrets, then the env var. Never hardcode."""
    try:
        return st.secrets["postgres"]["dsn"]
    except Exception:  # noqa: BLE001 - secrets may be absent entirely
        return os.environ.get("DATABASE_URL")


@st.cache_resource(show_spinner=False)
def get_store(dsn: str) -> VariantStore:
    store = VariantStore(dsn)
    store.ensure_schema()
    return store


@st.cache_data(ttl=60, show_spinner=False)
def _list_variant_dicts(dsn: str, report_key: str) -> list[dict]:
    store = get_store(dsn)
    return [
        {"id": v.id, "name": v.name, "columns": list(v.columns)}
        for v in store.list_variants(report_key)
    ]


def render_variant_panel(report_key: str, all_columns: list[str], key_prefix: str) -> list[str]:
    all_columns = list(all_columns)
    st.markdown("**Variant**")

    dsn = _dsn()
    if not dsn:
        st.caption("ℹ️ Shared variants unavailable — database not configured. Showing **Standard**.")
        return all_columns

    try:
        store = get_store(dsn)
        variants = _list_variant_dicts(dsn, report_key)
    except Exception as exc:  # noqa: BLE001 - never let the DB break the dashboard
        st.caption(f"⚠️ Shared variants unavailable ({type(exc).__name__}). Showing **Standard**.")
        return all_columns

    by_name = {v["name"]: v for v in variants}
    options = [STANDARD_NAME] + list(by_name.keys())

    sel_key = f"{key_prefix}_sel"
    ms_key = f"{key_prefix}_cols"
    last_key = f"{key_prefix}_last_sel"
    pending_key = f"{key_prefix}_pending_sel"
    flash_key = f"{key_prefix}_flash"

    # one-shot flash from a previous action
    if flash_key in st.session_state:
        kind, msg = st.session_state.pop(flash_key)
        getattr(st, kind, st.info)(msg)

    # apply a queued selection (from create/rename/delete) BEFORE the selectbox
    if pending_key in st.session_state:
        st.session_state[sel_key] = st.session_state.pop(pending_key)
    if st.session_state.get(sel_key) not in options:
        st.session_state[sel_key] = STANDARD_NAME

    selected = st.selectbox("Variant", options, key=sel_key, label_visibility="collapsed")

    # columns of the selected variant (Standard = all)
    if selected == STANDARD_NAME:
        base_cols = all_columns
    else:
        saved = by_name.get(selected, {}).get("columns", all_columns)
        base_cols = [c for c in saved if c in all_columns] or all_columns

    # reseed the multiselect when the variant changes (set BEFORE the widget)
    if ms_key not in st.session_state or st.session_state.get(last_key) != selected:
        st.session_state[last_key] = selected
        st.session_state[ms_key] = base_cols

    chosen = st.multiselect(
        "Columns shown", options=all_columns, key=ms_key, label_visibility="collapsed",
    )
    display_cols = chosen or all_columns  # never render an empty table
    is_standard = selected == STANDARD_NAME
    dirty = list(chosen) != list(base_cols)

    if dirty and is_standard:
        st.caption("● Custom columns — use **Save as new variant** to keep them.")
    elif dirty:
        st.caption("● Unsaved changes — click **Save**, or **Save as new** to keep them.")

    b_save, b_del = st.columns(2)
    if b_save.button("💾 Save", disabled=is_standard or not dirty,
                     key=f"{key_prefix}_btn_save", use_container_width=True):
        try:
            store.update_columns(by_name[selected]["id"], chosen)
            _list_variant_dicts.clear()
            st.session_state[flash_key] = ("success", f"Saved '{selected}'.")
        except VariantError as exc:
            st.session_state[flash_key] = ("error", str(exc))
        st.rerun()

    if b_del.button("🗑️ Delete", disabled=is_standard,
                    key=f"{key_prefix}_btn_del", use_container_width=True):
        try:
            store.delete_variant(by_name[selected]["id"])
            _list_variant_dicts.clear()
            st.session_state[pending_key] = STANDARD_NAME
            st.session_state[flash_key] = ("success", f"Deleted '{selected}'.")
        except VariantError as exc:
            st.session_state[flash_key] = ("error", str(exc))
        st.rerun()

    with st.expander("📑 Save as new variant"):
        new_name = st.text_input("New variant name", key=f"{key_prefix}_newname")
        if st.button("Create", type="primary", key=f"{key_prefix}_btn_create"):
            try:
                created = store.create_variant(report_key, new_name, chosen)
                _list_variant_dicts.clear()
                st.session_state[pending_key] = created.name
                st.session_state[flash_key] = ("success", f"Created '{created.name}'.")
                st.rerun()
            except VariantError as exc:
                st.error(str(exc))

    if not is_standard:
        with st.expander("✏️ Rename variant"):
            rn = st.text_input("New name", value=selected, key=f"{key_prefix}_rnname")
            if st.button("Rename", key=f"{key_prefix}_btn_rename"):
                try:
                    renamed = store.rename_variant(by_name[selected]["id"], rn)
                    _list_variant_dicts.clear()
                    st.session_state[pending_key] = renamed.name
                    st.session_state[flash_key] = ("success", f"Renamed to '{renamed.name}'.")
                    st.rerun()
                except VariantError as exc:
                    st.error(str(exc))

    return display_cols
