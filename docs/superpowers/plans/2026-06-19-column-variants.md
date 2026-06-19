# Shared Column Variants Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an SAP-style "Variant" picker above the Shortage Report (Delivery Fill Rate) and Unconfirmed Demand Report (Sales Order Fill Rate) tables that lets any user create/view/update/rename/delete shared, Postgres-persisted column layouts controlling both the on-screen table and its Excel sheet.

**Architecture:** A pure logic + Postgres data layer (`src/column_variants.py`, fully unit-tested) and a thin Streamlit panel (`src/column_variants_ui.py`). Pages 2 & 3 call the panel and project the chosen columns onto the displayed frame and the matching Excel sheet via `apply_columns(...)` — the engines are untouched. "Standard" (all columns) is built in and never stored.

**Tech Stack:** Python 3.14, Streamlit, pandas, **psycopg 3** (`psycopg[binary]`), hosted **Neon** Postgres, pytest.

**Spec:** `docs/superpowers/specs/2026-06-19-column-variants-design.md`

---

## File map

- **Create** `src/column_variants.py` — constants, `Variant` dataclass, validation helpers, `apply_columns`, `VariantStore` (CRUD + schema).
- **Create** `src/column_variants_ui.py` — `render_variant_panel(...)` Streamlit panel + DSN/cache helpers.
- **Create** `tests/test_column_variants.py` — pure-logic tests (always run) + DB-integration tests (gated on `TEST_DATABASE_URL`).
- **Create** `.gitignore`, `.streamlit/secrets.toml.example`.
- **Modify** `requirements.txt` — add `psycopg[binary]`.
- **Modify** `pages/2_Delivery_Fill_Rate_Dashboard.py` — imports, panel in Shortage Report tab, project Excel sheet.
- **Modify** `pages/3_Sales_Order_Fill_Rate_Dashboard.py` — imports, panel in Unconfirmed Demand tab, project Excel sheet.

## Environment notes (read once)

- Python 3.14.5 is on PATH. Run tests from the repo root: `python -m pytest tests/ -q`.
- `src/` and `tests/` are packages (`__init__.py` present); imports use `from src.x import y`.
- Install deps before starting: `python -m pip install -r requirements-dev.txt` (after Task 1 adds psycopg).
- Commit style in this repo is direct commits; this plan commits per task. Do **not** commit `.streamlit/secrets.toml` (Task 1 git-ignores it).

---

### Task 1: Dependencies, .gitignore, secrets scaffolding

**Files:**
- Modify: `requirements.txt`
- Create: `.gitignore`
- Create: `.streamlit/secrets.toml.example`

- [ ] **Step 1: Add psycopg to `requirements.txt`**

Final contents of `requirements.txt`:

```
pandas>=2.0
openpyxl>=3.1
streamlit>=1.30
plotly>=5.0
psycopg[binary]>=3.1
```

- [ ] **Step 2: Create `.gitignore`** (the repo currently has none)

```
__pycache__/
*.py[cod]
*.db
.venv/
venv/
.env
.streamlit/secrets.toml
```

- [ ] **Step 3: Create `.streamlit/secrets.toml.example`**

```toml
# Copy this file to .streamlit/secrets.toml (git-ignored) and fill in the DSN.
# On Streamlit Community Cloud, set the same [postgres] dsn in the app's Secrets UI.
[postgres]
# Neon POOLED connection string (host contains "-pooler"), sslmode=require.
dsn = "postgresql://USER:PASSWORD@ep-xxxx-pooler.REGION.aws.neon.tech/neondb?sslmode=require"
```

- [ ] **Step 4: Install deps**

Run: `python -m pip install -r requirements-dev.txt`
Expected: psycopg installs without error.

- [ ] **Step 5 (optional): stop tracking bytecode already in the index**

Run: `git rm -r --cached __pycache__ src/__pycache__ pages/__pycache__ tests/__pycache__`
Expected: removes tracked `.pyc` files from the index (working files stay). Skip if it reports no matches.

- [ ] **Step 6: Commit**

```bash
git add requirements.txt .gitignore .streamlit/secrets.toml.example
git commit -m "chore: add psycopg dep, .gitignore, and secrets example for column variants"
```

---

### Task 2: Provision Neon Postgres + local secret

**Files:**
- Create (local only, git-ignored): `.streamlit/secrets.toml`

> This task uses the Neon integration (MCP tools `create_project`, `get_connection_string`) or the Neon web console. It produces a connection string; nothing here is committed.

- [ ] **Step 1: Create a Neon project**

Create a new Neon project named `supplier-compliance-variants` (free tier). Default database `neondb` is fine.

- [ ] **Step 2: Get the POOLED connection string**

Capture the **pooled** connection string (host contains `-pooler`), with `?sslmode=require`. This is the DSN.

- [ ] **Step 3: Write the local secret**

Create `.streamlit/secrets.toml` (NOT committed):

```toml
[postgres]
dsn = "postgresql://USER:PASSWORD@ep-xxxx-pooler.REGION.aws.neon.tech/neondb?sslmode=require"
```

- [ ] **Step 4: Smoke-test connectivity**

Run:
```bash
python -c "import tomllib,psycopg; d=tomllib.load(open('.streamlit/secrets.toml','rb'))['postgres']['dsn']; c=psycopg.connect(d); print(c.execute('select 1').fetchone()); c.close()"
```
Expected: prints `(1,)`. (No commit — the secret is git-ignored.)

> Keep this DSN; you will also set it as `TEST_DATABASE_URL` to run the DB tests in Task 4, and paste it into Streamlit Cloud Secrets in Task 8.

---

### Task 3: `column_variants.py` — constants, model, pure helpers (TDD)

**Files:**
- Create: `src/column_variants.py`
- Test: `tests/test_column_variants.py`

- [ ] **Step 1: Write the failing pure-logic tests**

Create `tests/test_column_variants.py`:

```python
"""Tests for shared column variants."""
from __future__ import annotations

import pandas as pd
import pytest

from src.column_variants import (
    MAX_NAME_LEN,
    REPORT_DELIVERY_SHORTAGE,
    REPORT_SALES_ORDER_UNCONFIRMED,
    STANDARD_NAME,
    VALID_REPORT_KEYS,
    VariantValidationError,
    apply_columns,
    is_reserved_name,
    normalize_columns,
    validate_name,
    validate_report_key,
)


# -- validate_name -----------------------------------------------------------
def test_validate_name_trims_and_returns():
    assert validate_name("  My View  ") == "My View"


def test_validate_name_rejects_empty():
    with pytest.raises(VariantValidationError):
        validate_name("   ")


def test_validate_name_rejects_too_long():
    with pytest.raises(VariantValidationError):
        validate_name("x" * (MAX_NAME_LEN + 1))


def test_validate_name_rejects_reserved_any_case():
    with pytest.raises(VariantValidationError):
        validate_name("standard")
    with pytest.raises(VariantValidationError):
        validate_name("  STANDARD ")


# -- is_reserved_name --------------------------------------------------------
def test_is_reserved_name_case_insensitive():
    assert is_reserved_name("Standard")
    assert is_reserved_name(" sTaNdArD ")
    assert not is_reserved_name("My View")


# -- validate_report_key -----------------------------------------------------
def test_validate_report_key_accepts_known():
    assert validate_report_key(REPORT_DELIVERY_SHORTAGE) == REPORT_DELIVERY_SHORTAGE
    assert REPORT_SALES_ORDER_UNCONFIRMED in VALID_REPORT_KEYS


def test_validate_report_key_rejects_unknown():
    with pytest.raises(VariantValidationError):
        validate_report_key("nope")


# -- normalize_columns -------------------------------------------------------
def test_normalize_columns_dedupes_preserving_order():
    assert normalize_columns(["a", "b", "a", " c "]) == ["a", "b", "c"]


def test_normalize_columns_drops_blanks():
    assert normalize_columns(["a", "", "  ", "b"]) == ["a", "b"]


def test_normalize_columns_rejects_empty_result():
    with pytest.raises(VariantValidationError):
        normalize_columns([])
    with pytest.raises(VariantValidationError):
        normalize_columns(["", "   "])


def test_normalize_columns_rejects_non_list():
    with pytest.raises(VariantValidationError):
        normalize_columns("abc")


# -- apply_columns -----------------------------------------------------------
def _df():
    return pd.DataFrame({"a": [1], "b": [2], "c": [3]})


def test_apply_columns_projects_and_orders():
    out = apply_columns(_df(), ["c", "a"])
    assert list(out.columns) == ["c", "a"]


def test_apply_columns_skips_missing():
    out = apply_columns(_df(), ["c", "zzz", "a"])
    assert list(out.columns) == ["c", "a"]


def test_apply_columns_none_present_returns_unchanged():
    out = apply_columns(_df(), ["x", "y"])
    assert list(out.columns) == ["a", "b", "c"]


def test_apply_columns_empty_or_none_returns_unchanged():
    assert list(apply_columns(_df(), []).columns) == ["a", "b", "c"]
    assert list(apply_columns(_df(), None).columns) == ["a", "b", "c"]


def test_standard_name_constant():
    assert STANDARD_NAME == "Standard"
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest tests/test_column_variants.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'src.column_variants'`.

- [ ] **Step 3: Create `src/column_variants.py` with constants, model, and pure helpers**

```python
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
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_column_variants.py -q`
Expected: PASS (all pure-logic tests green).

- [ ] **Step 5: Commit**

```bash
git add src/column_variants.py tests/test_column_variants.py
git commit -m "feat: column-variant constants, model, and pure validation helpers"
```

---

### Task 4: `column_variants.py` — `VariantStore` (Postgres CRUD) + DB tests (TDD)

**Files:**
- Modify: `src/column_variants.py`
- Modify: `tests/test_column_variants.py`

- [ ] **Step 1: Append DB-integration tests (gated on `TEST_DATABASE_URL`)**

Add to the top imports of `tests/test_column_variants.py`:

```python
import os
import time
```

Append to `tests/test_column_variants.py`:

```python
from src.column_variants import (
    DuplicateVariantError,
    VariantNotFoundError,
    VariantStore,
)

TEST_DSN = os.environ.get("TEST_DATABASE_URL")
needs_db = pytest.mark.skipif(not TEST_DSN, reason="TEST_DATABASE_URL not set")
PREFIX = "pytest_cv_"


@pytest.fixture
def store():
    s = VariantStore(TEST_DSN)
    s.ensure_schema()
    yield s
    with s._connect() as conn:  # noqa: SLF001 - test cleanup
        conn.execute("DELETE FROM column_variants WHERE name LIKE %s", (PREFIX + "%",))


@needs_db
def test_create_list_get_roundtrip(store):
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "A", ["product", "plant"])
    assert v.id > 0
    assert v.columns == ["product", "plant"]
    fetched = store.get_variant(v.id)
    assert fetched.name == PREFIX + "A"
    names = [x.name for x in store.list_variants(REPORT_DELIVERY_SHORTAGE)]
    assert PREFIX + "A" in names


@needs_db
def test_duplicate_name_case_insensitive(store):
    store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Dup", ["product"])
    with pytest.raises(DuplicateVariantError):
        store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "dup", ["plant"])


@needs_db
def test_update_columns_bumps_updated_at(store):
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Upd", ["product"])
    time.sleep(0.02)
    v2 = store.update_columns(v.id, ["plant", "short_amount"])
    assert v2.columns == ["plant", "short_amount"]
    assert v2.updated_at > v2.created_at


@needs_db
def test_rename(store):
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Old", ["product"])
    v2 = store.rename_variant(v.id, PREFIX + "New")
    assert v2.name == PREFIX + "New"


@needs_db
def test_rename_into_existing_name_rejected(store):
    store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Taken", ["product"])
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Move", ["plant"])
    with pytest.raises(DuplicateVariantError):
        store.rename_variant(v.id, PREFIX + "taken")


@needs_db
def test_delete(store):
    v = store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "Gone", ["product"])
    store.delete_variant(v.id)
    with pytest.raises(VariantNotFoundError):
        store.get_variant(v.id)


@needs_db
def test_variants_separated_by_report_key(store):
    store.create_variant(REPORT_DELIVERY_SHORTAGE, PREFIX + "DelOnly", ["product"])
    so_names = [x.name for x in store.list_variants(REPORT_SALES_ORDER_UNCONFIRMED)]
    assert PREFIX + "DelOnly" not in so_names
```

- [ ] **Step 2: Run DB tests to verify they fail (with a DB configured)**

Set the env var to the Neon DSN, then run:

PowerShell: `$env:TEST_DATABASE_URL = (python -c "import tomllib;print(tomllib.load(open('.streamlit/secrets.toml','rb'))['postgres']['dsn'])"); python -m pytest tests/test_column_variants.py -q`

Expected: the `needs_db` tests FAIL with `ImportError`/`AttributeError` (no `VariantStore`). (If `TEST_DATABASE_URL` is unset they SKIP — set it so they actually run.)

- [ ] **Step 3: Append `VariantStore` and schema to `src/column_variants.py`**

Add these imports at the top of `src/column_variants.py` (below `import pandas as pd`):

```python
import psycopg
from psycopg.errors import UniqueViolation
from psycopg.rows import dict_row
from psycopg.types.json import Jsonb
```

Append to the end of `src/column_variants.py`:

```python
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
            conn.execute("DELETE FROM column_variants WHERE id = %s", (variant_id,))
```

- [ ] **Step 4: Run the full test file to verify it passes**

Run (with `TEST_DATABASE_URL` set as in Step 2): `python -m pytest tests/test_column_variants.py -q`
Expected: PASS — pure-logic + all `needs_db` tests green.

- [ ] **Step 5: Commit**

```bash
git add src/column_variants.py tests/test_column_variants.py
git commit -m "feat: Postgres-backed VariantStore with CRUD + DB integration tests"
```

---

### Task 5: `column_variants_ui.py` — Streamlit panel

**Files:**
- Create: `src/column_variants_ui.py`
- Test: `tests/test_column_variants_ui_import.py`

- [ ] **Step 1: Write a failing import smoke test**

Create `tests/test_column_variants_ui_import.py`:

```python
"""Smoke test: the UI module imports and exposes the panel entrypoint."""
def test_ui_module_imports():
    import src.column_variants_ui as ui
    assert callable(ui.render_variant_panel)
    assert callable(ui.get_store)
```

- [ ] **Step 2: Run it to verify it fails**

Run: `python -m pytest tests/test_column_variants_ui_import.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'src.column_variants_ui'`.

- [ ] **Step 3: Create `src/column_variants_ui.py`**

```python
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
```

- [ ] **Step 4: Run the import test to verify it passes**

Run: `python -m pytest tests/test_column_variants_ui_import.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/column_variants_ui.py tests/test_column_variants_ui_import.py
git commit -m "feat: Streamlit Variant panel (DB-backed, Standard fallback)"
```

---

### Task 6: Wire the Delivery Fill Rate Shortage Report (page 2)

**Files:**
- Modify: `pages/2_Delivery_Fill_Rate_Dashboard.py` (imports near top; `tab_short` ~lines 366-374; download call ~line 517)

- [ ] **Step 1: Add imports**

Find the existing import block that includes `from src.fill_rate_engine import (...)`. Immediately after that import statement, add:

```python
from src.column_variants import REPORT_DELIVERY_SHORTAGE, apply_columns
from src.column_variants_ui import render_variant_panel
```

- [ ] **Step 2: Add the panel + projection in the Shortage Report tab**

Replace this block (currently ~lines 366-374):

```python
# ── Shortage Report ─────────────────────────────────────────────────────────
with tab_short:
    st.subheader(f"Shortage Report — {len(shortage_df):,} problem lines")
    st.caption("Sorted by highest Short Amount. Use sidebar filters to narrow results.")

    if shortage_df.empty:
        st.success("No shortages or fill rate issues found.")
    else:
        st.dataframe(shortage_df, use_container_width=True, hide_index=True)
```

with:

```python
# ── Shortage Report ─────────────────────────────────────────────────────────
with tab_short:
    st.subheader(f"Shortage Report — {len(shortage_df):,} problem lines")
    st.caption("Sorted by highest Short Amount. Use sidebar filters to narrow results.")

    shortage_cols = render_variant_panel(
        REPORT_DELIVERY_SHORTAGE,
        list(shortage_df.columns),
        key_prefix="dfr_shortage_variant",
    )

    if shortage_df.empty:
        st.success("No shortages or fill rate issues found.")
    else:
        st.dataframe(
            apply_columns(shortage_df, shortage_cols),
            use_container_width=True,
            hide_index=True,
        )
```

- [ ] **Step 3: Project the Excel sheet**

In the download tab, change the `shortage_df=shortage_df,` argument (currently ~line 517) inside the `generate_excel_report(` call to:

```python
                shortage_df=apply_columns(shortage_df, shortage_cols),
```

- [ ] **Step 4: Verify the app runs and the feature works**

Run: `python -m streamlit run app.py`
Then in the browser:
1. Open **Delivery Fill Rate Dashboard**, upload a delivery export.
2. Go to the **Shortage Report** tab. Expect a **Variant** picker (dropdown = `Standard`) above the grid and a "Columns shown" multiselect listing all columns.
3. Deselect a couple of columns → the grid updates immediately.
4. **Save as new variant** → name `Mgmt View` → it appears in the dropdown and is selected.
5. Switch dropdown `Standard` ↔ `Mgmt View` → columns change accordingly.
6. **Refresh the browser** → `Mgmt View` still listed (persisted in Postgres).
7. Go to **Download Report**, generate the workbook, open the **Shortage Report** sheet → only the variant's columns appear; other sheets unchanged.

Expected: all of the above; no exceptions in the terminal.

- [ ] **Step 5: Commit**

```bash
git add pages/2_Delivery_Fill_Rate_Dashboard.py
git commit -m "feat: column-variant picker on Delivery Shortage Report (table + Excel)"
```

---

### Task 7: Wire the Sales Order Unconfirmed Demand Report (page 3)

**Files:**
- Modify: `pages/3_Sales_Order_Fill_Rate_Dashboard.py` (imports near top; `tab_unc` ~lines 366-395; download call ~line 529)

- [ ] **Step 1: Add imports**

Find the existing import block that includes `from src.sales_order_engine import (...)`. Immediately after it, add:

```python
from src.column_variants import REPORT_SALES_ORDER_UNCONFIRMED, apply_columns
from src.column_variants_ui import render_variant_panel
```

- [ ] **Step 2: Add the panel in the Unconfirmed Demand tab**

Replace the start of the `tab_unc` block (currently ~lines 366-371):

```python
# ── Unconfirmed Demand Report ─────────────────────────────────────────────────
with tab_unc:
    st.subheader(f"Unconfirmed Demand Report — {len(unconfirmed_df):,} problem lines")

    if unconfirmed_df.empty:
        st.success("No unconfirmed demand found in this dataset.")
```

with (insert the panel call so `unconf_cols` is always defined, even when empty):

```python
# ── Unconfirmed Demand Report ─────────────────────────────────────────────────
with tab_unc:
    st.subheader(f"Unconfirmed Demand Report — {len(unconfirmed_df):,} problem lines")

    unconf_cols = render_variant_panel(
        REPORT_SALES_ORDER_UNCONFIRMED,
        list(unconfirmed_df.columns),
        key_prefix="sov_unconfirmed_variant",
    )

    if unconfirmed_df.empty:
        st.success("No unconfirmed demand found in this dataset.")
```

- [ ] **Step 3: Project the displayed table**

Within the same `else:` branch, the table is rendered (currently ~line 395):

```python
        st.dataframe(view, use_container_width=True, hide_index=True)
```

Change it to:

```python
        st.dataframe(apply_columns(view, unconf_cols), use_container_width=True, hide_index=True)
```

- [ ] **Step 4: Project the Excel sheet**

In the download tab, change the `unconfirmed_df=unconfirmed_df,` argument (currently ~line 529) inside the `generate_excel_report(` call to:

```python
                unconfirmed_df=apply_columns(unconfirmed_df, unconf_cols),
```

- [ ] **Step 5: Verify the app runs and the feature works**

Run: `python -m streamlit run app.py`
Then in the browser:
1. Open **Sales Order Fill Rate Dashboard**, upload a sales-order export.
2. Go to the **Unconfirmed Demand Report** tab. Expect the **Variant** picker above the existing local filters.
3. Deselect columns → the grid updates (the local Key Account/Plant/etc. filters still work for rows).
4. **Save as new variant** `SO Slim`; refresh the browser → it persists.
5. Confirm it is independent from the Delivery dashboard's variants (different dropdown contents).
6. Download the report → the **Unconfirmed Demand Rpt** sheet has only the variant's columns; other sheets unchanged.

Expected: all of the above; no terminal exceptions.

- [ ] **Step 6: Commit**

```bash
git add pages/3_Sales_Order_Fill_Rate_Dashboard.py
git commit -m "feat: column-variant picker on Sales Order Unconfirmed Demand (table + Excel)"
```

---

### Task 8: Full regression + deploy

**Files:** none (verification + deploy)

- [ ] **Step 1: Run the whole test suite**

Run (with `TEST_DATABASE_URL` set to the Neon DSN): `python -m pytest tests/ -q`
Expected: PASS — existing tests (billback, compliance, po_exclusion) plus the new column-variant tests.

- [ ] **Step 2: DB-down fallback check**

Temporarily rename `.streamlit/secrets.toml` (or unset `DATABASE_URL`) and run the app: the Variant panel should show "Shared variants unavailable … Showing **Standard**" and the tables/Excel must still work. Restore the secret afterward.

- [ ] **Step 3: Set the secret on Streamlit Cloud**

In the deployed app's **Settings → Secrets**, add:

```toml
[postgres]
dsn = "postgresql://USER:PASSWORD@ep-xxxx-pooler.REGION.aws.neon.tech/neondb?sslmode=require"
```

(This step is manual and must be done by the app owner — it cannot be committed.)

- [ ] **Step 4: Push and verify on Cloud**

Push the branch/commits to `origin/main` (per this repo's deploy flow, Streamlit Cloud redeploys from `origin/main`). After redeploy, repeat the Task 6/7 browser checks on the live app and confirm a variant created on Cloud **survives a manual reboot/redeploy**.

---

## Self-review (completed by plan author)

**Spec coverage:**
- Standard read-only built-in → Task 3 (`STANDARD_NAME`, `apply_columns` passthrough) + Task 5 (dropdown, disabled buttons). ✅
- CRUD by all users, no auth → Task 4 (`VariantStore`), Task 5 (panel buttons). ✅
- Persist across Cloud restarts/redeploys → Tasks 2, 8 (Neon + Cloud secret). ✅
- Separated by report key → `report_key` column + `validate_report_key` (Task 4); distinct `REPORT_*` constants wired per page (Tasks 6, 7). ✅
- Active variant drives table **and** Excel → Tasks 6 & 7 (`apply_columns` on display frame and on the `generate_excel_report` argument). ✅
- Case-insensitive unique names → `ux_column_variants_report_name_ci` + `DuplicateVariantError` (Task 4); reserved "Standard" (Task 3). ✅
- `updated_at = now()` on update/rename → Task 4 SQL. ✅
- DB CHECKs on report_key + non-empty columns array → Task 4 schema; `normalize_columns` rejects empty (Task 3). ✅
- DB-down graceful fallback → Task 5 try/except; verified Task 8 Step 2. ✅
- No localStorage/JSON; psycopg only new dep → Task 1. ✅
- `.gitignore` + secrets via Streamlit secrets/env → Tasks 1, 2, 8. ✅
- Engines untouched → Tasks 6 & 7 project before passing. ✅

**Placeholder scan:** none — every step has concrete code/commands.

**Type consistency:** `render_variant_panel(report_key, all_columns, key_prefix)`, `apply_columns(df, columns)`, `VariantStore` method names, and the `REPORT_*` / `STANDARD_NAME` constants are used identically across Tasks 3–7. ✅
