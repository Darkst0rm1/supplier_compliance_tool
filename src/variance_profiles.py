"""Saved Variance Profiles — data layer.

Lets users save the filter sets they commonly use on the Delivery Variance and
Sales Order Variance dashboards, so they don't rebuild the same filters every
visit. Profiles are stored in a local SQLite database.

Security / architecture notes (adapted for a local single-user Streamlit tool):

* Identity is NEVER taken from the client/UI. ``get_context()`` derives the
  tenant and user from the local OS session (env-overridable tenant + the OS
  login name). Every query is scoped by ``tenant_id`` and ``user_id``.
* ``require_permission()`` is a real gate that currently grants the local user
  every permission, structured so a future auth layer can replace it.
* ``run_action()`` is the single choke-point for every mutation: it checks the
  permission, runs the change inside one transaction, and writes an audit-log
  row. This mirrors the spec's ``runAction()`` / ``requirePermission()``.
* ``validate_filters()`` is the Python analogue of the requested Zod schema —
  it validates and normalises the filter JSON, raising on bad shapes/types.
* Filters are stored as validated JSON in the ``filters`` column.

The public name of the feature is **Saved Variance Profiles**; the dashboard
surfaces it as **Variance Profile**.
"""
from __future__ import annotations

import getpass
import json
import os
import sqlite3
import threading
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

# ---------------------------------------------------------------------------
# Profile types (the requested enum)
# ---------------------------------------------------------------------------
PROFILE_TYPE_DELIVERY = "DELIVERY"
PROFILE_TYPE_SALES_ORDER = "SALES_ORDER"
PROFILE_TYPES = (PROFILE_TYPE_DELIVERY, PROFILE_TYPE_SALES_ORDER)

# Audit actions
ACTION_CREATE = "CREATE"
ACTION_UPDATE = "UPDATE"
ACTION_DUPLICATE = "DUPLICATE"
ACTION_DELETE = "DELETE"
ACTION_SET_DEFAULT = "SET_DEFAULT"
ACTION_RESET_DEFAULT = "RESET_DEFAULT"


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------
class ProfileError(Exception):
    """Base error for the profiles feature."""


class ProfileValidationError(ProfileError):
    """Raised when a name, profile type, or filter payload is invalid."""


class ProfileNotFoundError(ProfileError):
    """Raised when a profile id does not exist for the current tenant/user."""


class SystemProfileError(ProfileError):
    """Raised when a caller tries to edit or delete a system profile."""


class PermissionDeniedError(ProfileError):
    """Raised by require_permission when the context lacks a permission."""


# ---------------------------------------------------------------------------
# Context — tenant + user, derived from the local session, never the client
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Context:
    tenant_id: str
    user_id: str
    # permissions the context holds; the local user holds all of them
    permissions: frozenset[str] = frozenset({"*"})


def get_context() -> Context:
    """Return the tenant/user context from the local environment.

    The tenant defaults to ``local`` and may be overridden with the
    ``SCT_TENANT_ID`` environment variable. The user is the OS login name
    (override with ``SCT_USER_ID``). Nothing here is read from UI input.
    """
    tenant_id = os.environ.get("SCT_TENANT_ID", "local").strip() or "local"
    user_id = os.environ.get("SCT_USER_ID", "").strip()
    if not user_id:
        try:
            user_id = getpass.getuser()
        except Exception:
            user_id = "local-user"
    return Context(tenant_id=tenant_id, user_id=user_id)


# permission required for each mutating action
_ACTION_PERMISSIONS: dict[str, str] = {
    ACTION_CREATE: "variance_profile:write",
    ACTION_UPDATE: "variance_profile:write",
    ACTION_DUPLICATE: "variance_profile:write",
    ACTION_DELETE: "variance_profile:delete",
    ACTION_SET_DEFAULT: "variance_profile:write",
    ACTION_RESET_DEFAULT: "variance_profile:write",
}


def require_permission(ctx: Context, permission: str) -> None:
    """Permission gate. The local user holds the wildcard, so this grants
    everything today; a real auth layer can swap the check without callers
    changing."""
    if "*" in ctx.permissions or permission in ctx.permissions:
        return
    raise PermissionDeniedError(f"Context {ctx.user_id} lacks permission: {permission}")


# ---------------------------------------------------------------------------
# Filter schema validation (Zod analogue)
# ---------------------------------------------------------------------------
_LIST_FIELDS = (
    "suppliers",
    "customers",
    "warehouses",            # plants / DCs
    "delivery_numbers",
    "sales_order_numbers",
    "products",              # products / materials
    "product_groups",
    "variance_types",
    "statuses",
    "visible_columns",
)
_FLOAT_FIELDS = ("variance_amount_threshold", "variance_pct_threshold")
_NULLABLE_STR_FIELDS = ("sort_field", "date_start", "date_end")

DEFAULT_FILTERS: dict[str, Any] = {
    **{f: [] for f in _LIST_FIELDS},
    **{f: None for f in _FLOAT_FIELDS},
    **{f: None for f in _NULLABLE_STR_FIELDS},
    "date_preference": "all",      # "all" | "custom"
    "sort_direction": "desc",      # "asc" | "desc"
    "extra": {},                   # any other active filters
}


def default_filters() -> dict[str, Any]:
    """Return a fresh copy of the canonical empty filter set."""
    return json.loads(json.dumps(DEFAULT_FILTERS))


def validate_filters(raw: Any) -> dict[str, Any]:
    """Validate and normalise a filter payload.

    Returns a new dict containing every canonical key. Known keys are
    type-checked; unrecognised keys are preserved under ``extra`` so the schema
    can evolve without losing data. Raises ProfileValidationError on bad types.
    """
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ProfileValidationError(f"filters is not valid JSON: {exc}") from exc
    if raw is None:
        raw = {}
    if not isinstance(raw, dict):
        raise ProfileValidationError("filters must be a JSON object")

    out = default_filters()
    extra = dict(out["extra"])

    for key, value in raw.items():
        if key in _LIST_FIELDS:
            if value is None:
                value = []
            if not isinstance(value, (list, tuple)):
                raise ProfileValidationError(f"'{key}' must be a list")
            out[key] = [str(v) for v in value]
        elif key in _FLOAT_FIELDS:
            if value in (None, ""):
                out[key] = None
            else:
                try:
                    out[key] = float(value)
                except (TypeError, ValueError) as exc:
                    raise ProfileValidationError(f"'{key}' must be a number or null") from exc
        elif key in _NULLABLE_STR_FIELDS:
            out[key] = None if value in (None, "") else str(value)
        elif key == "date_preference":
            if value not in ("all", "custom"):
                raise ProfileValidationError("'date_preference' must be 'all' or 'custom'")
            out[key] = value
        elif key == "sort_direction":
            if value not in ("asc", "desc"):
                raise ProfileValidationError("'sort_direction' must be 'asc' or 'desc'")
            out[key] = value
        elif key == "extra":
            if value is None:
                value = {}
            if not isinstance(value, dict):
                raise ProfileValidationError("'extra' must be an object")
            extra.update(value)
        else:
            # unknown top-level key -> keep it under extra
            extra[key] = value

    out["extra"] = extra
    return out


def validate_name(name: Any) -> str:
    if not isinstance(name, str):
        raise ProfileValidationError("name must be a string")
    name = name.strip()
    if not name:
        raise ProfileValidationError("name cannot be empty")
    if len(name) > 120:
        raise ProfileValidationError("name cannot exceed 120 characters")
    return name


def validate_profile_type(profile_type: Any) -> str:
    if profile_type not in PROFILE_TYPES:
        raise ProfileValidationError(
            f"profile_type must be one of {PROFILE_TYPES}, got {profile_type!r}"
        )
    return profile_type


# ---------------------------------------------------------------------------
# System profiles (seeded, read-only, duplicable)
# ---------------------------------------------------------------------------
def _f(**overrides: Any) -> dict[str, Any]:
    return validate_filters(overrides)


SYSTEM_PROFILES: dict[str, list[tuple[str, dict[str, Any]]]] = {
    PROFILE_TYPE_DELIVERY: [
        ("All Delivery Variances",
         _f(variance_amount_threshold=1000.0)),
        ("Missing Deliveries",
         _f(variance_amount_threshold=1000.0, statuses=["Shorted"],
            sort_field="Highest Short Amount")),
        ("Quantity Variance",
         _f(variance_amount_threshold=1000.0, statuses=["Shorted"],
            sort_field="Highest Short Qty")),
        ("Date Variance",
         _f(variance_amount_threshold=1000.0, sort_field="Requested Delivery Date",
            sort_direction="asc")),
        ("Status Variance",
         _f(variance_amount_threshold=1000.0,
            wh_fill_status=["WH Fill Rate Issue"], customer_fill_status=["Customer Impact"])),
    ],
    PROFILE_TYPE_SALES_ORDER: [
        ("All Sales Order Variances",
         _f(variance_amount_threshold=500.0)),
        ("Missing Sales Orders",
         _f(variance_amount_threshold=500.0, statuses=["Unconfirmed Demand"],
            sort_field="Highest Unconfirmed Amount")),
        ("Quantity Variance",
         _f(variance_amount_threshold=500.0, statuses=["Unconfirmed Demand"],
            sort_field="Highest Unconfirmed Qty")),
        ("Value Variance",
         _f(variance_amount_threshold=500.0, sort_field="Highest Unconfirmed Amount")),
        ("Status Variance",
         _f(variance_amount_threshold=500.0, variance_types=["High Priority", "Medium Priority"])),
    ],
}

# The canonical "show everything" profile per type — the system default and the
# target of "reset to system default".
SYSTEM_DEFAULT_NAME: dict[str, str] = {
    PROFILE_TYPE_DELIVERY: "All Delivery Variances",
    PROFILE_TYPE_SALES_ORDER: "All Sales Order Variances",
}


# ---------------------------------------------------------------------------
# Profile record
# ---------------------------------------------------------------------------
@dataclass
class Profile:
    id: int
    tenant_id: str
    user_id: str
    name: str
    profile_type: str
    filters: dict[str, Any]
    is_default: bool
    is_system: bool
    created_at: str
    updated_at: str

    @property
    def editable(self) -> bool:
        return not self.is_system


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _row_to_profile(row: sqlite3.Row) -> Profile:
    return Profile(
        id=row["id"],
        tenant_id=row["tenant_id"],
        user_id=row["user_id"],
        name=row["name"],
        profile_type=row["profile_type"],
        filters=validate_filters(row["filters"]),
        is_default=bool(row["is_default"]),
        is_system=bool(row["is_system"]),
        created_at=row["created_at"],
        updated_at=row["updated_at"],
    )


# ---------------------------------------------------------------------------
# Store
# ---------------------------------------------------------------------------
_DEFAULT_DB_PATH = Path(__file__).resolve().parent.parent / "data" / "variance_profiles.db"


class ProfileStore:
    """Tenant- and user-scoped store for Saved Variance Profiles."""

    def __init__(self, db_path: str | os.PathLike | None = None, ctx: Context | None = None):
        self.db_path = str(db_path) if db_path is not None else str(_DEFAULT_DB_PATH)
        if self.db_path != ":memory:":
            Path(self.db_path).parent.mkdir(parents=True, exist_ok=True)
        self.ctx = ctx or get_context()
        self._lock = threading.Lock()
        self._conn = sqlite3.connect(self.db_path, check_same_thread=False)
        self._conn.row_factory = sqlite3.Row
        self._conn.execute("PRAGMA foreign_keys = ON")
        self._init_db()
        self._seed_system_profiles()

    # -- schema -------------------------------------------------------------
    def _init_db(self) -> None:
        with self._conn:
            self._conn.execute(
                """
                CREATE TABLE IF NOT EXISTS variance_profiles (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    tenant_id    TEXT NOT NULL,
                    user_id      TEXT NOT NULL,
                    name         TEXT NOT NULL,
                    profile_type TEXT NOT NULL CHECK (profile_type IN ('DELIVERY','SALES_ORDER')),
                    filters      TEXT NOT NULL,
                    is_default   INTEGER NOT NULL DEFAULT 0,
                    is_system    INTEGER NOT NULL DEFAULT 0,
                    created_at   TEXT NOT NULL,
                    updated_at   TEXT NOT NULL,
                    UNIQUE (tenant_id, user_id, profile_type, name)
                )
                """
            )
            self._conn.execute(
                """
                CREATE TABLE IF NOT EXISTS variance_profile_audit (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    tenant_id    TEXT NOT NULL,
                    user_id      TEXT NOT NULL,
                    profile_id   INTEGER,
                    profile_type TEXT,
                    name         TEXT,
                    action       TEXT NOT NULL,
                    detail       TEXT,
                    created_at   TEXT NOT NULL
                )
                """
            )
            self._conn.execute(
                "CREATE INDEX IF NOT EXISTS ix_profiles_scope "
                "ON variance_profiles (tenant_id, user_id, profile_type)"
            )

    # -- seeding ------------------------------------------------------------
    def _seed_system_profiles(self) -> None:
        """Ensure the read-only system profiles exist for this tenant/user.

        System profiles are seeded per tenant+user so each user gets the same
        starting set; they are flagged ``is_system`` and cannot be edited or
        deleted.
        """
        with self._lock, self._conn:
            for profile_type, profiles in SYSTEM_PROFILES.items():
                for name, filters in profiles:
                    exists = self._conn.execute(
                        "SELECT 1 FROM variance_profiles "
                        "WHERE tenant_id=? AND user_id=? AND profile_type=? AND name=? AND is_system=1",
                        (self.ctx.tenant_id, self.ctx.user_id, profile_type, name),
                    ).fetchone()
                    if exists:
                        continue
                    now = _now()
                    self._conn.execute(
                        "INSERT INTO variance_profiles "
                        "(tenant_id, user_id, name, profile_type, filters, is_default, is_system, created_at, updated_at) "
                        "VALUES (?,?,?,?,?,0,1,?,?)",
                        (
                            self.ctx.tenant_id,
                            self.ctx.user_id,
                            name,
                            profile_type,
                            json.dumps(validate_filters(filters)),
                            now,
                            now,
                        ),
                    )

    # -- audit + action wrapper --------------------------------------------
    def _write_audit(
        self,
        action: str,
        profile_id: int | None,
        profile_type: str | None,
        name: str | None,
        detail: dict[str, Any] | None = None,
    ) -> None:
        self._conn.execute(
            "INSERT INTO variance_profile_audit "
            "(tenant_id, user_id, profile_id, profile_type, name, action, detail, created_at) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (
                self.ctx.tenant_id,
                self.ctx.user_id,
                profile_id,
                profile_type,
                name,
                action,
                json.dumps(detail) if detail is not None else None,
                _now(),
            ),
        )

    def run_action(self, action: str, fn: Callable[[sqlite3.Connection], Any]) -> Any:
        """Single choke-point for mutations: permission check + transaction +
        audit log. ``fn`` receives the live connection and must return a tuple
        ``(result, audit_kwargs)`` where ``audit_kwargs`` feeds ``_write_audit``.
        """
        permission = _ACTION_PERMISSIONS.get(action, "variance_profile:write")
        require_permission(self.ctx, permission)
        with self._lock, self._conn:  # transaction
            result, audit = fn(self._conn)
            self._write_audit(action=action, **audit)
        return result

    # -- queries (tenant + user scoped) ------------------------------------
    def list_profiles(self, profile_type: str) -> list[Profile]:
        validate_profile_type(profile_type)
        rows = self._conn.execute(
            "SELECT * FROM variance_profiles "
            "WHERE tenant_id=? AND user_id=? AND profile_type=? "
            "ORDER BY is_system DESC, name COLLATE NOCASE ASC",
            (self.ctx.tenant_id, self.ctx.user_id, profile_type),
        ).fetchall()
        return [_row_to_profile(r) for r in rows]

    def get_profile(self, profile_id: int) -> Profile:
        row = self._conn.execute(
            "SELECT * FROM variance_profiles WHERE id=? AND tenant_id=? AND user_id=?",
            (profile_id, self.ctx.tenant_id, self.ctx.user_id),
        ).fetchone()
        if row is None:
            raise ProfileNotFoundError(f"profile {profile_id} not found")
        return _row_to_profile(row)

    def get_default_profile(self, profile_type: str) -> Profile:
        """Return the user's default profile for the type, or the system
        default ("All ...") when no explicit default is set."""
        validate_profile_type(profile_type)
        row = self._conn.execute(
            "SELECT * FROM variance_profiles "
            "WHERE tenant_id=? AND user_id=? AND profile_type=? AND is_default=1",
            (self.ctx.tenant_id, self.ctx.user_id, profile_type),
        ).fetchone()
        if row is not None:
            return _row_to_profile(row)
        # fall back to the canonical system default
        row = self._conn.execute(
            "SELECT * FROM variance_profiles "
            "WHERE tenant_id=? AND user_id=? AND profile_type=? AND name=? AND is_system=1",
            (self.ctx.tenant_id, self.ctx.user_id, profile_type, SYSTEM_DEFAULT_NAME[profile_type]),
        ).fetchone()
        if row is None:
            raise ProfileNotFoundError(f"no default profile for {profile_type}")
        return _row_to_profile(row)

    # -- mutations ----------------------------------------------------------
    def create_profile(self, name: str, profile_type: str, filters: Any) -> Profile:
        name = validate_name(name)
        validate_profile_type(profile_type)
        clean = validate_filters(filters)

        def _do(conn: sqlite3.Connection):
            now = _now()
            try:
                cur = conn.execute(
                    "INSERT INTO variance_profiles "
                    "(tenant_id, user_id, name, profile_type, filters, is_default, is_system, created_at, updated_at) "
                    "VALUES (?,?,?,?,?,0,0,?,?)",
                    (self.ctx.tenant_id, self.ctx.user_id, name, profile_type,
                     json.dumps(clean), now, now),
                )
            except sqlite3.IntegrityError as exc:
                raise ProfileValidationError(
                    f"a {profile_type} profile named '{name}' already exists"
                ) from exc
            new_id = int(cur.lastrowid)
            audit = {"profile_id": new_id, "profile_type": profile_type, "name": name}
            return new_id, audit

        new_id = self.run_action(ACTION_CREATE, _do)
        return self.get_profile(new_id)

    def update_profile(self, profile_id: int, *, filters: Any = None, name: str | None = None) -> Profile:
        existing = self.get_profile(profile_id)
        if existing.is_system:
            raise SystemProfileError("system profiles cannot be edited")
        new_name = existing.name if name is None else validate_name(name)
        new_filters = existing.filters if filters is None else validate_filters(filters)

        def _do(conn: sqlite3.Connection):
            try:
                conn.execute(
                    "UPDATE variance_profiles SET name=?, filters=?, updated_at=? "
                    "WHERE id=? AND tenant_id=? AND user_id=?",
                    (new_name, json.dumps(new_filters), _now(),
                     profile_id, self.ctx.tenant_id, self.ctx.user_id),
                )
            except sqlite3.IntegrityError as exc:
                raise ProfileValidationError(
                    f"a {existing.profile_type} profile named '{new_name}' already exists"
                ) from exc
            audit = {"profile_id": profile_id, "profile_type": existing.profile_type, "name": new_name}
            return profile_id, audit

        self.run_action(ACTION_UPDATE, _do)
        return self.get_profile(profile_id)

    def duplicate_profile(self, profile_id: int, new_name: str | None = None) -> Profile:
        src = self.get_profile(profile_id)
        target_name = validate_name(new_name) if new_name else self._unique_copy_name(src)

        def _do(conn: sqlite3.Connection):
            now = _now()
            try:
                cur = conn.execute(
                    "INSERT INTO variance_profiles "
                    "(tenant_id, user_id, name, profile_type, filters, is_default, is_system, created_at, updated_at) "
                    "VALUES (?,?,?,?,?,0,0,?,?)",
                    (self.ctx.tenant_id, self.ctx.user_id, target_name, src.profile_type,
                     json.dumps(src.filters), now, now),
                )
            except sqlite3.IntegrityError as exc:
                raise ProfileValidationError(
                    f"a {src.profile_type} profile named '{target_name}' already exists"
                ) from exc
            new_id = int(cur.lastrowid)
            audit = {
                "profile_id": new_id,
                "profile_type": src.profile_type,
                "name": target_name,
                "detail": {"duplicated_from": profile_id, "source_name": src.name},
            }
            return new_id, audit

        new_id = self.run_action(ACTION_DUPLICATE, _do)
        return self.get_profile(new_id)

    def _unique_copy_name(self, src: Profile) -> str:
        base = f"{src.name} (copy)"
        candidate = base
        i = 2
        existing = {p.name for p in self.list_profiles(src.profile_type)}
        while candidate in existing:
            candidate = f"{base} {i}"
            i += 1
        return candidate

    def delete_profile(self, profile_id: int) -> None:
        existing = self.get_profile(profile_id)
        if existing.is_system:
            raise SystemProfileError("system profiles cannot be deleted")

        def _do(conn: sqlite3.Connection):
            conn.execute(
                "DELETE FROM variance_profiles WHERE id=? AND tenant_id=? AND user_id=?",
                (profile_id, self.ctx.tenant_id, self.ctx.user_id),
            )
            audit = {"profile_id": profile_id, "profile_type": existing.profile_type, "name": existing.name}
            return None, audit

        self.run_action(ACTION_DELETE, _do)

    def set_default(self, profile_id: int) -> Profile:
        existing = self.get_profile(profile_id)

        def _do(conn: sqlite3.Connection):
            conn.execute(
                "UPDATE variance_profiles SET is_default=0, updated_at=updated_at "
                "WHERE tenant_id=? AND user_id=? AND profile_type=?",
                (self.ctx.tenant_id, self.ctx.user_id, existing.profile_type),
            )
            conn.execute(
                "UPDATE variance_profiles SET is_default=1 "
                "WHERE id=? AND tenant_id=? AND user_id=?",
                (profile_id, self.ctx.tenant_id, self.ctx.user_id),
            )
            audit = {"profile_id": profile_id, "profile_type": existing.profile_type, "name": existing.name}
            return profile_id, audit

        self.run_action(ACTION_SET_DEFAULT, _do)
        return self.get_profile(profile_id)

    def reset_default(self, profile_type: str) -> Profile:
        """Clear the user's default for a type so the system default applies."""
        validate_profile_type(profile_type)

        def _do(conn: sqlite3.Connection):
            conn.execute(
                "UPDATE variance_profiles SET is_default=0 "
                "WHERE tenant_id=? AND user_id=? AND profile_type=?",
                (self.ctx.tenant_id, self.ctx.user_id, profile_type),
            )
            audit = {"profile_id": None, "profile_type": profile_type, "name": SYSTEM_DEFAULT_NAME[profile_type]}
            return None, audit

        self.run_action(ACTION_RESET_DEFAULT, _do)
        return self.get_default_profile(profile_type)

    # -- audit access (for tests / inspection) -----------------------------
    def list_audit(self, limit: int = 100) -> list[dict[str, Any]]:
        rows = self._conn.execute(
            "SELECT * FROM variance_profile_audit "
            "WHERE tenant_id=? AND user_id=? ORDER BY id DESC LIMIT ?",
            (self.ctx.tenant_id, self.ctx.user_id, limit),
        ).fetchall()
        return [dict(r) for r in rows]

    def close(self) -> None:
        self._conn.close()
