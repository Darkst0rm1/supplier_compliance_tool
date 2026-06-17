"""Saved Variance Profiles — Streamlit UI.

Renders the **Variance Profile** panel (dropdown + actions) used by the
Delivery Variance and Sales Order Variance dashboards, and the glue that maps a
page's filter widgets to/from the canonical filter payload stored by
``variance_profiles``.

Design (works within Streamlit's top-to-bottom rerun model):

* Each page declares a list of :class:`FieldSpec` describing every filter widget
  and the canonical filter key it maps to. The same spec drives both *seeding*
  (profile -> widgets) and *collecting* (widgets -> profile).
* Widget keys are versioned (``<prefix>_<key>_<version>``). Loading a profile
  bumps the version and seeds the new keys, which both applies the saved values
  and clears any stale widget state.
* The panel reads "current filters" from ``st.session_state`` (last run's widget
  values), so it can be rendered above the widgets and still detect unsaved
  changes accurately.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Callable

import streamlit as st

from src.variance_profiles import (
    Profile,
    ProfileError,
    ProfileStore,
    SystemProfileError,
    default_filters,
    validate_filters,
)


# ---------------------------------------------------------------------------
# Field specification
# ---------------------------------------------------------------------------
@dataclass
class FieldSpec:
    """Maps one filter widget to a canonical filter key.

    ``path`` is the canonical key; use ``"extra.<name>"`` for page-specific
    filters that have no dedicated canonical field. ``kind`` is one of
    ``multiselect``, ``select``, ``number``, ``daterange``, ``columns``.
    ``options`` returns the current choices (and, for ``daterange``, a
    ``(min_date, max_date)`` tuple).
    """
    path: str
    key: str
    kind: str
    options: Callable[[], Any] | None = None
    default: Any = None   # fallback for number fields when the profile value is None


# ---------------------------------------------------------------------------
# Store accessor (one shared connection per process)
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_store() -> ProfileStore:
    return ProfileStore()


# ---------------------------------------------------------------------------
# Dotted-path helpers for the canonical filter dict
# ---------------------------------------------------------------------------
def _get_path(filters: dict[str, Any], path: str) -> Any:
    if path.startswith("extra."):
        return filters.get("extra", {}).get(path[len("extra."):])
    return filters.get(path)


def _set_path(filters: dict[str, Any], path: str, value: Any) -> None:
    if path.startswith("extra."):
        filters.setdefault("extra", {})[path[len("extra."):]] = value
    else:
        filters[path] = value


def versioned_key(prefix: str, key: str, version: int) -> str:
    return f"{prefix}_{key}_{version}"


# ---------------------------------------------------------------------------
# Seed (profile -> widget session_state) and collect (widgets -> filters)
# ---------------------------------------------------------------------------
def seed_session_state(prefix: str, version: int, specs: list[FieldSpec], filters: dict[str, Any]) -> None:
    """Write widget values into session_state from a profile's filters.

    Values are intersected with the currently available options so Streamlit
    never errors on a saved value that isn't in the current dataset.
    """
    for spec in specs:
        wkey = versioned_key(prefix, spec.key, version)
        value = _get_path(filters, spec.path)
        opts = spec.options() if spec.options else None

        if spec.kind in ("multiselect",):
            available = list(opts or [])
            chosen = [v for v in (value or []) if v in available]
            st.session_state[wkey] = chosen
        elif spec.kind == "columns":
            available = list(opts or [])
            if value:  # explicit subset
                st.session_state[wkey] = [v for v in value if v in available]
            else:       # empty == show all
                st.session_state[wkey] = list(available)
        elif spec.kind == "select":
            available = list(opts or [])
            st.session_state[wkey] = value if value in available else (available[0] if available else None)
        elif spec.kind == "number":
            if value not in (None, ""):
                st.session_state[wkey] = float(value)
            else:
                st.session_state[wkey] = float(spec.default) if spec.default is not None else 0.0
        elif spec.kind == "daterange":
            min_d, max_d = (opts or (None, None))
            start = _parse_date(_get_path(filters, "date_start"))
            end = _parse_date(_get_path(filters, "date_end"))
            pref = filters.get("date_preference", "all")
            if pref == "custom" and start and end and min_d and max_d:
                start = max(start, min_d)
                end = min(end, max_d)
                st.session_state[wkey] = (start, end)
            elif min_d and max_d:
                st.session_state[wkey] = (min_d, max_d)


def collect_filters(
    prefix: str,
    version: int,
    specs: list[FieldSpec],
    base_filters: dict[str, Any],
) -> dict[str, Any]:
    """Build a canonical filter dict from current widget values.

    Starts from ``base_filters`` (the active profile) so any canonical fields
    that this page doesn't expose as widgets are preserved untouched.
    """
    out = validate_filters(base_filters)
    for spec in specs:
        wkey = versioned_key(prefix, spec.key, version)
        if wkey not in st.session_state:
            continue
        raw = st.session_state[wkey]
        if spec.kind in ("multiselect",):
            _set_path(out, spec.path, [str(v) for v in (raw or [])])
        elif spec.kind == "columns":
            available = list(spec.options() if spec.options else [])
            chosen = [str(v) for v in (raw or [])]
            # empty / "everything selected" both mean "all" -> store []
            _set_path(out, spec.path, [] if not chosen or set(chosen) == set(available) else chosen)
        elif spec.kind == "select":
            _set_path(out, spec.path, raw if raw not in ("", None) else None)
        elif spec.kind == "number":
            _set_path(out, spec.path, float(raw) if raw not in ("", None) else None)
        elif spec.kind == "daterange":
            min_d, max_d = (spec.options() if spec.options else (None, None))
            if isinstance(raw, (tuple, list)) and len(raw) == 2 and all(raw):
                start, end = raw
                if min_d and max_d and start == min_d and end == max_d:
                    out["date_preference"] = "all"
                    out["date_start"] = out["date_end"] = None
                else:
                    out["date_preference"] = "custom"
                    out["date_start"] = start.isoformat()
                    out["date_end"] = end.isoformat()
    return out


def _parse_date(value: Any) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def filters_equal(a: dict[str, Any], b: dict[str, Any]) -> bool:
    return json.dumps(validate_filters(a), sort_keys=True) == json.dumps(validate_filters(b), sort_keys=True)


# ---------------------------------------------------------------------------
# Session-state keys for the panel itself
# ---------------------------------------------------------------------------
def _ss(prefix: str, name: str) -> str:
    return f"{prefix}__vp_{name}"


def _request_load(prefix: str, profile_id: int) -> None:
    """Queue a profile to be loaded on the next run and bump the widget
    version so the filter widgets get re-seeded from it."""
    st.session_state[_ss(prefix, "pending_load")] = profile_id
    st.session_state[_ss(prefix, "version")] = st.session_state.get(_ss(prefix, "version"), 0) + 1


# ---------------------------------------------------------------------------
# Main panel
# ---------------------------------------------------------------------------
def render_variance_profile_panel(
    store: ProfileStore,
    profile_type: str,
    prefix: str,
    specs: list[FieldSpec],
    container=None,
) -> tuple[Profile, int]:
    """Render the Variance Profile dropdown + actions.

    Returns ``(active_profile, version)``. The caller then builds its filter
    widgets using ``versioned_key(prefix, spec.key, version)``; the widgets pick
    up the seeded session_state automatically.
    """
    ui = container or st
    ver_key = _ss(prefix, "version")
    active_key = _ss(prefix, "active_id")
    pending_key = _ss(prefix, "pending_load")

    # First run: queue the default profile so widgets start seeded.
    if active_key not in st.session_state:
        default_profile = store.get_default_profile(profile_type)
        st.session_state[active_key] = default_profile.id
        st.session_state[ver_key] = 0
        st.session_state[pending_key] = default_profile.id

    version = st.session_state.get(ver_key, 0)

    # Apply a queued load: seed widget state from the target profile.
    pending_id = st.session_state.pop(pending_key, None)
    if pending_id is not None:
        try:
            target = store.get_profile(pending_id)
        except ProfileError:
            target = store.get_default_profile(profile_type)
        st.session_state[active_key] = target.id
        seed_session_state(prefix, version, specs, target.filters)

    # Resolve the active profile (may have been deleted out from under us).
    try:
        active = store.get_profile(st.session_state[active_key])
    except ProfileError:
        active = store.get_default_profile(profile_type)
        st.session_state[active_key] = active.id

    profiles = store.list_profiles(profile_type)
    current_filters = collect_filters(prefix, version, specs, active.filters)
    dirty = not filters_equal(current_filters, active.filters)

    # ---- header ----------------------------------------------------------
    ui.markdown("#### Variance Profile")

    def _label(p: Profile) -> str:
        tag = "🔒 " if p.is_system else ""
        star = " ⭐" if p.is_default else ""
        return f"{tag}{p.name}{star}"

    ids = [p.id for p in profiles]
    labels = {p.id: _label(p) for p in profiles}
    try:
        idx = ids.index(active.id)
    except ValueError:
        idx = 0

    chosen_id = ui.selectbox(
        "Variance Profile",
        ids,
        index=idx,
        format_func=lambda i: labels.get(i, str(i)),
        key=_ss(prefix, "select"),
        label_visibility="collapsed",
    )
    if chosen_id != active.id:
        _request_load(prefix, chosen_id)
        st.rerun()

    # status line
    bits = []
    bits.append("System profile (read-only)" if active.is_system else "Your profile")
    if active.is_default:
        bits.append("default")
    ui.caption(" · ".join(bits))
    if dirty:
        ui.warning("● Unsaved changes — click **Update Profile** to save.", icon="⚠️")
    else:
        ui.caption("✓ No unsaved changes.")

    # ---- actions ---------------------------------------------------------
    msg_key = _ss(prefix, "flash")
    if msg_key in st.session_state:
        kind, text = st.session_state.pop(msg_key)
        getattr(ui, kind, ui.info)(text)

    def _flash(kind: str, text: str) -> None:
        st.session_state[msg_key] = (kind, text)

    # Update / Save-as-new row
    c1, c2 = ui.columns(2)
    update_disabled = active.is_system or not dirty
    if c1.button("💾 Update Profile", key=_ss(prefix, "btn_update"),
                 disabled=update_disabled, use_container_width=True,
                 help="Save current filters into this profile" if not active.is_system
                      else "System profiles can't be edited — duplicate it first"):
        try:
            store.update_profile(active.id, filters=current_filters)
            _flash("success", f"Updated '{active.name}'.")
        except ProfileError as exc:
            _flash("error", str(exc))
        st.rerun()

    if c2.button("⭐ Set as Default", key=_ss(prefix, "btn_default"),
                 disabled=active.is_default, use_container_width=True):
        try:
            store.set_default(active.id)
            _flash("success", f"'{active.name}' is now your default {profile_type.replace('_', ' ').lower()} profile.")
        except ProfileError as exc:
            _flash("error", str(exc))
        st.rerun()

    c3, c4 = ui.columns(2)
    if c3.button("📑 Duplicate", key=_ss(prefix, "btn_dup"), use_container_width=True,
                 help="Create an editable copy of this profile"):
        try:
            dup = store.duplicate_profile(active.id)
            _flash("success", f"Created '{dup.name}'.")
            _request_load(prefix, dup.id)
        except ProfileError as exc:
            _flash("error", str(exc))
        st.rerun()

    delete_disabled = active.is_system
    if c4.button("🗑️ Delete", key=_ss(prefix, "btn_del"),
                 disabled=delete_disabled, use_container_width=True,
                 help="Delete this profile" if not active.is_system else "System profiles can't be deleted"):
        try:
            store.delete_profile(active.id)
            _flash("success", f"Deleted '{active.name}'.")
            _request_load(prefix, store.get_default_profile(profile_type).id)
        except ProfileError as exc:
            _flash("error", str(exc))
        st.rerun()

    # Save as new (name input + button)
    with ui.expander("➕ Save current filters as a new profile"):
        new_name = st.text_input("Profile name", key=_ss(prefix, "new_name"),
                                  placeholder="e.g. My Delivery Variance")
        if st.button("Save New Profile", key=_ss(prefix, "btn_savenew"), type="primary"):
            try:
                created = store.create_profile(new_name, profile_type, current_filters)
                _flash("success", f"Saved '{created.name}'.")
                st.session_state[_ss(prefix, "new_name")] = ""
                _request_load(prefix, created.id)
                st.rerun()
            except ProfileError as exc:
                ui.error(str(exc))

    if ui.button("↩️ Reset to System Default", key=_ss(prefix, "btn_reset"),
                 use_container_width=True):
        try:
            sysdef = store.reset_default(profile_type)
            _flash("success", "Reset to the system default profile.")
            _request_load(prefix, sysdef.id)
        except ProfileError as exc:
            _flash("error", str(exc))
        st.rerun()

    return active, version
