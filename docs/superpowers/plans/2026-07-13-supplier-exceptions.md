# Supplier Exceptions Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an `Exception Status` column to the Supplier Summary sheet and a "Should Have Uploaded" sheet listing suppliers who uploaded nothing despite being expected to, backed by a Neon-hosted exceptions list seeded from the Master Inbound Delivery Compliance Tracker.

**Architecture:** A new `src/supplier_exceptions.py` store mirrors the existing `src/column_variants.py` Postgres pattern (psycopg3, short-lived autocommit connections, fail-open when the DB is unreachable). A new `src/tracker_importer.py` reads the exceptions out of the tracker workbook. `compliance_engine.build_report` gains an optional `exceptions` argument; when omitted, every output sheet is byte-identical to today.

**Tech Stack:** Python 3.14, pandas 3.0.3, Streamlit, psycopg3 (`psycopg[binary]`), pytest, openpyxl.

## Global Constraints

- **Bill-back and the compliance percentage MUST NOT change.** Exception status is informational only. Task 5 includes a regression test that enforces this; do not weaken it.
- `build_report(...)` must remain callable **without** `exceptions` and produce identical output to today. Existing tests and callers must not be edited to accommodate the new parameter.
- **Fail open, never crash.** If the Postgres DSN is absent or the DB is unreachable, the report still generates, the Exception column reads `Expected to upload` for everyone, and the page shows an `st.info` explaining why. Nobody is ever silently marked exempt by a failure.
- Never hardcode a DSN. Read `st.secrets["postgres"]["dsn"]`, falling back to the `DATABASE_URL` env var — exactly as `src/column_variants_ui.py:21-26` does.
- pandas 3.0 empty-frame rules apply: never call `Series.map()` on an empty lookup or `DataFrame.apply(axis=1)` on an empty frame. Use the existing `_map_lookup` / `_apply_rows` helpers in `compliance_engine.py`.
- No page file may call `st.set_page_config` — it lives only in `app.py`.
- Python is not on PATH. Use `C:\Users\melgh\AppData\Local\Python\pythoncore-3.14-64\python.exe`.
- Run the full suite with: `python -m pytest tests/ -q`

## File Structure

| File | Responsibility |
|---|---|
| `src/normalizer.py` (modify) | Gains `normalize_supplier_name` — the name join key. Lives here because it's the sibling of the existing `normalize_po`. |
| `src/config.py` (modify) | The three `Exception Status` label constants + the tracker's literal sheet/column/status strings. |
| `src/tracker_importer.py` (create) | Reads the tracker workbook → list of `(supplier_name, reason)`. Knows nothing about Postgres. |
| `src/supplier_exceptions.py` (create) | `ExceptionRecord` dataclass, `ExceptionStore` (Postgres), and the pure `classify_supplier()` function. Knows nothing about Streamlit or Excel. |
| `src/compliance_engine.py` (modify) | Annotates SAP rows with `Exception Status`; new column on Supplier Summary; new `Should Have Uploaded` sheet. |
| `pages/1_Supplier_Compliance_Dashboard.py` (modify) | Loads exceptions, passes them to `build_report`, renders the new section + a management expander. |
| `scripts/seed_supplier_exceptions.py` (create) | Idempotent one-off seed from the tracker workbook. |
| `tests/test_supplier_exceptions.py` (create) | All tests for the above. |

---

### Task 1: Supplier-name normalization

**Files:**
- Modify: `src/normalizer.py` (append at end)
- Test: `tests/test_supplier_exceptions.py` (create)

**Interfaces:**
- Produces: `normalize_supplier_name(value) -> str` — used by every later task as the join key between the tracker and SAP.

- [ ] **Step 1: Write the failing test**

Create `tests/test_supplier_exceptions.py`:

```python
"""Tests for the supplier exceptions feature."""
from __future__ import annotations

import pytest

from src.normalizer import normalize_supplier_name


class TestNormalizeSupplierName:
    def test_uppercases_and_strips(self):
        assert normalize_supplier_name("  Acetum S.P.A.  ") == "ACETUM S P A"

    def test_collapses_internal_whitespace(self):
        assert normalize_supplier_name("BOTHWELL    CHEESE") == "BOTHWELL CHEESE"

    def test_punctuation_becomes_space(self):
        # The tracker and SAP punctuate differently; both must land on the same key.
        assert normalize_supplier_name("DARE (LESLEY STOWE FINE FOODS)") == (
            "DARE LESLEY STOWE FINE FOODS"
        )
        assert normalize_supplier_name("D&D ITALIA SPA") == "D D ITALIA SPA"
        assert normalize_supplier_name("C.H. GUENTHER & SON, INC.") == (
            "C H GUENTHER SON INC"
        )

    def test_sap_and_tracker_spellings_converge(self):
        # Real pair from the June export vs the tracker.
        assert normalize_supplier_name("SERFUNGHI DI CALABRETTA LUIGI") == (
            normalize_supplier_name("Serfunghi di Calabretta Luigi")
        )

    def test_empty_and_missing(self):
        assert normalize_supplier_name(None) == ""
        assert normalize_supplier_name("") == ""
        assert normalize_supplier_name("   ") == ""
        assert normalize_supplier_name(float("nan")) == ""
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: FAIL — `ImportError: cannot import name 'normalize_supplier_name'`

- [ ] **Step 3: Write minimal implementation**

Append to `src/normalizer.py`:

```python
# Characters that differ freely between the tracker's spelling of a supplier and
# SAP's. Mapped to a space (not removed) so "D&D" -> "D D", never "DD".
_SUPPLIER_PUNCT = ".,'-()&/"


def normalize_supplier_name(value) -> str:
    """Normalize a supplier/vendor name into a join key. Empty/NaN -> ''.

    The tracker identifies suppliers only by name -- its `Supplier #` column
    (G61, 491) has zero overlap with SAP's 8-digit Vendor Number -- so this is
    the only key available. Uppercases, maps punctuation to spaces, collapses
    whitespace.
    """
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    for ch in _SUPPLIER_PUNCT:
        s = s.replace(ch, " ")
    return _WHITESPACE_RE.sub(" ", s).strip().upper()
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: PASS (5 passed)

- [ ] **Step 5: Commit**

```bash
git add src/normalizer.py tests/test_supplier_exceptions.py
git commit -m "feat: add normalize_supplier_name join key"
```

---

### Task 2: Tracker importer

**Files:**
- Create: `src/tracker_importer.py`
- Modify: `src/config.py` (append)
- Test: `tests/test_supplier_exceptions.py` (append)

**Interfaces:**
- Consumes: `normalize_supplier_name` from Task 1.
- Produces: `read_tracker_exceptions(path_or_buffer) -> list[tuple[str, str]]` returning `(supplier_name, reason)` pairs, de-duplicated by normalized name; `TrackerImportError`.

- [ ] **Step 1: Append the config constants**

Append to `src/config.py`:

```python
# --- Supplier exceptions ---------------------------------------------------
# Suppliers with an approved exception are not required to upload inbound
# documentation. The list is sourced from the Master Inbound Delivery Compliance
# Tracker workbook, as the union of two lists:
#   1. Tracker sheet, Compliance Status == TRACKER_STATUS_UNABLE_TO_COMPLY (24)
#   2. "POs received" sheet, a column literally containing "EXEMPT" (3)
# The two overlap on 2 suppliers, so the union is 25.
TRACKER_SHEET = "Tracker"
TRACKER_STATUS_COLUMN = "Compliance Status"
TRACKER_NAME_COLUMN = "Supplier Names "  # NB: trailing space, as in the workbook

# NB: DOUBLE space after the dash. This is the literal value in the workbook --
# do not "correct" it. The Summary sheet words it differently ("NO - Unable to
# comply - Approved exceptions"); that sheet is not the source of truth.
TRACKER_STATUS_UNABLE_TO_COMPLY = "NO -  Unable to Comply"

TRACKER_EXEMPT_SHEET = "POs received"
TRACKER_EXEMPT_MARKER = "EXEMPT"

REASON_UNABLE_TO_COMPLY = "Unable to Comply"
REASON_EXEMPT_MARK = "EXEMPT mark"
REASON_MANUAL = "Manual"

# Exception Status values shown on the Supplier Summary sheet.
EXCEPTION_STATUS_EXCEPTION = "Exception"
EXCEPTION_STATUS_EXPECTED = "Expected to upload"
EXCEPTION_STATUS_NOT_ON_TRACKER = "Not on tracker"
```

**Deviation from the spec, deliberate:** the spec said a DB outage should make the column read `Unknown (DB unavailable)`. Instead, an outage falls back to `Expected to upload` for everyone plus an `st.info` banner on the page. Reason: `Unknown` would be a fourth state that exists only during a failure, and every consumer (the summary column, the chase-list filter) would need to decide what it means. Failing open to "expected" is the conservative reading — nobody is silently marked exempt — and the banner tells the user why the column is uniform. One less state, same safety.

- [ ] **Step 2: Write the failing test**

Append to `tests/test_supplier_exceptions.py`:

```python
import pandas as pd

from src.config import REASON_EXEMPT_MARK, REASON_UNABLE_TO_COMPLY
from src.tracker_importer import TrackerImportError, read_tracker_exceptions


def _fake_tracker(tmp_path):
    """A miniature stand-in for the real tracker workbook."""
    path = tmp_path / "tracker.xlsx"
    tracker = pd.DataFrame({
        "Supplier Names ": [
            "ACETUM S.P.A.",
            "BOTHWELL CHEESE",
            "DARE (LESLEY STOWE FINE FOODS)",
            "COMPLIANT CO",
        ],
        "Compliance Status": [
            "NO -  Unable to Comply",   # double space, as in the real workbook
            "NO -  Unable to Comply",
            "NO -  Unable to Comply",
            "YES - Submitted on Portal",
        ],
    })
    # "POs received": name in col 0, the EXEMPT marker in col 3.
    pos_received = pd.DataFrame({
        "Non compliant": ["BOTHWELL CHEESE", "LUNDBERG FAMILY FARMS", "OTHER CO"],
        "Pos received": [15, 12, 3],
        "Column1": ["x", "x", "x"],
        "Unnamed: 3": ["EXEMPT", "EXEMPT", None],
    })
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        tracker.to_excel(w, sheet_name="Tracker", index=False)
        pos_received.to_excel(w, sheet_name="POs received", index=False)
    return path


class TestReadTrackerExceptions:
    def test_unions_both_lists_and_dedupes(self, tmp_path):
        rows = read_tracker_exceptions(_fake_tracker(tmp_path))
        names = {n for n, _ in rows}
        # 3 Unable-to-Comply + 2 EXEMPT-marked, overlapping on Bothwell -> 4.
        assert names == {
            "ACETUM S.P.A.",
            "BOTHWELL CHEESE",
            "DARE (LESLEY STOWE FINE FOODS)",
            "LUNDBERG FAMILY FARMS",
        }
        assert "COMPLIANT CO" not in names
        assert "OTHER CO" not in names

    def test_unable_to_comply_wins_the_reason_on_overlap(self, tmp_path):
        rows = dict(read_tracker_exceptions(_fake_tracker(tmp_path)))
        assert rows["BOTHWELL CHEESE"] == REASON_UNABLE_TO_COMPLY
        assert rows["LUNDBERG FAMILY FARMS"] == REASON_EXEMPT_MARK

    def test_missing_sheet_raises_friendly_error(self, tmp_path):
        path = tmp_path / "wrong.xlsx"
        pd.DataFrame({"a": [1]}).to_excel(path, sheet_name="Nope", index=False)
        with pytest.raises(TrackerImportError, match="Tracker"):
            read_tracker_exceptions(path)
```

- [ ] **Step 3: Run test to verify it fails**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'src.tracker_importer'`

- [ ] **Step 4: Write the implementation**

Create `src/tracker_importer.py`:

```python
"""Read the approved-exception supplier list out of the Master Inbound Delivery
Compliance Tracker workbook.

The exceptions list is the UNION of two lists that Ops maintains by hand:
  1. Tracker sheet, Compliance Status == "NO -  Unable to Comply" (the Summary
     sheet calls these "Approved exceptions").
  2. "POs received" sheet, rows hand-marked "EXEMPT".

They overlap; "Unable to Comply" wins the reason because it is the
dropdown-driven, counted list.
"""
from __future__ import annotations

import pandas as pd

from .config import (
    REASON_EXEMPT_MARK,
    REASON_UNABLE_TO_COMPLY,
    TRACKER_EXEMPT_MARKER,
    TRACKER_EXEMPT_SHEET,
    TRACKER_NAME_COLUMN,
    TRACKER_SHEET,
    TRACKER_STATUS_COLUMN,
    TRACKER_STATUS_UNABLE_TO_COMPLY,
)
from .normalizer import normalize_supplier_name


class TrackerImportError(Exception):
    """The tracker workbook could not be read."""


def _find_column(df: pd.DataFrame, wanted: str) -> str | None:
    """Match a column tolerantly -- the workbook's headers carry stray spaces."""
    target = wanted.strip().casefold()
    for col in df.columns:
        if str(col).strip().casefold() == target:
            return col
    return None


def _unable_to_comply(path_or_buffer) -> list[str]:
    try:
        df = pd.read_excel(path_or_buffer, sheet_name=TRACKER_SHEET)
    except ValueError as exc:
        raise TrackerImportError(
            f"The workbook has no '{TRACKER_SHEET}' sheet. Is this the Master "
            "Inbound Delivery Compliance Tracker?"
        ) from exc

    name_col = _find_column(df, TRACKER_NAME_COLUMN)
    status_col = _find_column(df, TRACKER_STATUS_COLUMN)
    if name_col is None or status_col is None:
        raise TrackerImportError(
            f"The '{TRACKER_SHEET}' sheet needs both a "
            f"'{TRACKER_NAME_COLUMN.strip()}' and a '{TRACKER_STATUS_COLUMN}' column."
        )

    status = df[status_col].fillna("").astype(str).str.strip()
    hit = status == TRACKER_STATUS_UNABLE_TO_COMPLY.strip()
    return [str(n).strip() for n in df.loc[hit, name_col].dropna()]


def _exempt_marked(path_or_buffer) -> list[str]:
    """Names on any row of 'POs received' carrying an EXEMPT marker.

    The marker sits in an unnamed column with no stable header, so scan every
    column of the row rather than relying on a position.
    """
    try:
        df = pd.read_excel(path_or_buffer, sheet_name=TRACKER_EXEMPT_SHEET)
    except ValueError:
        return []  # This sheet is optional; the Tracker sheet is not.

    if df.empty or len(df.columns) < 2:
        return []

    name_col = df.columns[0]
    others = df.columns[1:]
    marked = pd.Series(False, index=df.index)
    for col in others:
        cells = df[col].fillna("").astype(str).str.strip().str.upper()
        marked |= cells == TRACKER_EXEMPT_MARKER

    return [str(n).strip() for n in df.loc[marked, name_col].dropna() if str(n).strip()]


def read_tracker_exceptions(path_or_buffer) -> list[tuple[str, str]]:
    """Return de-duplicated (supplier_name, reason) pairs from the tracker.

    De-duplication is by normalized name; the first spelling seen wins, and
    "Unable to Comply" wins the reason because it is scanned first.
    """
    out: list[tuple[str, str]] = []
    seen: set[str] = set()

    for names, reason in (
        (_unable_to_comply(path_or_buffer), REASON_UNABLE_TO_COMPLY),
        (_exempt_marked(path_or_buffer), REASON_EXEMPT_MARK),
    ):
        for name in names:
            key = normalize_supplier_name(name)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append((name, reason))
    return out
```

**Note on `path_or_buffer`:** `pd.read_excel` is called twice. A file *path* can be read twice safely; an uploaded file *buffer* cannot. The seed script (Task 7) passes a path, so this is fine — do not wire this function to a Streamlit uploader without seeking the buffer back to 0 between reads.

- [ ] **Step 5: Run test to verify it passes**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: PASS (8 passed)

- [ ] **Step 6: Verify against the REAL tracker**

Run:
```bash
python -c "from src.tracker_importer import read_tracker_exceptions; r = read_tracker_exceptions(r'C:\Users\melgh\Downloads\Master Inbound Delivery  Compliance Tracker_April 2025.xlsx'); print(len(r)); [print(' ', n, '|', why) for n, why in sorted(r)]"
```
Expected: `25`, with `LUNDBERG FAMILY FARMS | EXEMPT mark` and the other 24 as `Unable to Comply`. If this prints 24 or 27, the union/dedup logic is wrong — stop and fix.

- [ ] **Step 7: Commit**

```bash
git add src/config.py src/tracker_importer.py tests/test_supplier_exceptions.py
git commit -m "feat: read the approved-exception list from the tracker workbook"
```

---

### Task 3: Exceptions store + classification

**Files:**
- Create: `src/supplier_exceptions.py`
- Test: `tests/test_supplier_exceptions.py` (append)

**Interfaces:**
- Consumes: `normalize_supplier_name` (Task 1); the `EXCEPTION_STATUS_*` and `REASON_*` constants (Task 2).
- Produces:
  - `ExceptionRecord(id, supplier_name, normalized_name, vendor_number, reason, added_by, added_at)`
  - `ExceptionStore(dsn)` with `.ensure_schema()`, `.load_exceptions() -> dict[str, ExceptionRecord]`, `.add_exception(name, reason, vendor_number=None) -> ExceptionRecord`, `.remove_exception(normalized_name) -> None`
  - `classify_supplier(vendor_name, vendor_number, exceptions, tracker_names) -> str`
  - `ExceptionError`, `ExceptionValidationError`, `ExceptionNotFoundError`

**Design note — why `classify_supplier` takes `tracker_names`:** distinguishing `Expected to upload` from `Not on tracker` needs the set of *all* suppliers the tracker knows about, not just the exceptions. Pass an empty set to collapse the distinction (everything non-exception becomes `Expected to upload`).

- [ ] **Step 1: Write the failing test**

Append to `tests/test_supplier_exceptions.py`:

```python
from src.config import (
    EXCEPTION_STATUS_EXCEPTION,
    EXCEPTION_STATUS_EXPECTED,
    EXCEPTION_STATUS_NOT_ON_TRACKER,
)
from src.supplier_exceptions import (
    ExceptionRecord,
    ExceptionValidationError,
    classify_supplier,
    validate_supplier_name,
)


def _rec(name, reason="Unable to Comply", vendor_number=None):
    return ExceptionRecord(
        id=1,
        supplier_name=name,
        normalized_name=normalize_supplier_name(name),
        vendor_number=vendor_number,
        reason=reason,
        added_by=None,
        added_at=None,
    )


class TestClassifySupplier:
    def setup_method(self):
        self.exceptions = {
            normalize_supplier_name("BOTHWELL CHEESE"): _rec("BOTHWELL CHEESE"),
            normalize_supplier_name("CAFFE MAURO SPA"): _rec(
                "CAFFE MAURO SPA", vendor_number="70006979"
            ),
        }
        self.tracker = {
            normalize_supplier_name("BOTHWELL CHEESE"),
            normalize_supplier_name("CAFFE MAURO SPA"),
            normalize_supplier_name("ACQUA MINERALE SAN BENEDETTO"),
        }

    def test_exception_supplier(self):
        assert classify_supplier(
            "Bothwell Cheese", "70001111", self.exceptions, self.tracker
        ) == EXCEPTION_STATUS_EXCEPTION

    def test_vendor_number_matches_even_when_the_name_differs(self):
        # SAP spells it differently, but we recorded the vendor number.
        assert classify_supplier(
            "CAFFE MAURO S.P.A. (ITALY)", "70006979", self.exceptions, self.tracker
        ) == EXCEPTION_STATUS_EXCEPTION

    def test_on_tracker_but_not_an_exception(self):
        assert classify_supplier(
            "ACQUA MINERALE SAN BENEDETTO", "70007212", self.exceptions, self.tracker
        ) == EXCEPTION_STATUS_EXPECTED

    def test_absent_from_the_tracker(self):
        # A 3PL warehouse, not a supplier at all.
        assert classify_supplier(
            "AMERICOLD TACOMA", "70009999", self.exceptions, self.tracker
        ) == EXCEPTION_STATUS_NOT_ON_TRACKER

    def test_empty_tracker_collapses_to_expected(self):
        assert classify_supplier(
            "ANYONE", "70009999", self.exceptions, set()
        ) == EXCEPTION_STATUS_EXPECTED

    def test_blank_vendor_name(self):
        assert classify_supplier("", "", self.exceptions, self.tracker) == (
            EXCEPTION_STATUS_NOT_ON_TRACKER
        )


class TestValidateSupplierName:
    def test_rejects_blank(self):
        with pytest.raises(ExceptionValidationError):
            validate_supplier_name("   ")

    def test_trims(self):
        assert validate_supplier_name("  Acme  ") == "Acme"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'src.supplier_exceptions'`

- [ ] **Step 3: Write the implementation**

Create `src/supplier_exceptions.py`:

```python
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

import psycopg
from psycopg.errors import UniqueViolation
from psycopg.rows import dict_row

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
    return ExceptionRecord(
        id=row["id"],
        supplier_name=row["supplier_name"],
        normalized_name=row["normalized_name"],
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
        return psycopg.connect(self.dsn, autocommit=True, row_factory=dict_row)

    def ensure_schema(self) -> None:
        with self._connect() as conn:
            for stmt in _SCHEMA_STATEMENTS:
                conn.execute(stmt)

    def load_exceptions(self) -> dict[str, ExceptionRecord]:
        """All exceptions, keyed by normalized_name."""
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT {_COLS} FROM supplier_exceptions ORDER BY supplier_name"
            ).fetchall()
        return {r["normalized_name"]: _row_to_record(r) for r in rows}

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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: PASS (16 passed)

- [ ] **Step 5: Add the live-DB integration test**

These are gated on `TEST_DATABASE_URL`, exactly as the column-variants DB tests are — they skip by default. Append to `tests/test_supplier_exceptions.py`:

```python
import os

from src.supplier_exceptions import (
    DuplicateExceptionError,
    ExceptionNotFoundError,
    ExceptionStore,
)

_TEST_DSN = os.environ.get("TEST_DATABASE_URL")
requires_db = pytest.mark.skipif(not _TEST_DSN, reason="TEST_DATABASE_URL not set")


@requires_db
class TestExceptionStore:
    def setup_method(self):
        self.store = ExceptionStore(_TEST_DSN)
        self.store.ensure_schema()
        self.name = "PYTEST TEMP SUPPLIER"
        try:
            self.store.remove_exception(normalize_supplier_name(self.name))
        except ExceptionNotFoundError:
            pass

    def teardown_method(self):
        try:
            self.store.remove_exception(normalize_supplier_name(self.name))
        except ExceptionNotFoundError:
            pass

    def test_add_load_remove_roundtrip(self):
        rec = self.store.add_exception(self.name, REASON_MANUAL)
        assert rec.normalized_name == normalize_supplier_name(self.name)

        loaded = self.store.load_exceptions()
        assert rec.normalized_name in loaded
        assert loaded[rec.normalized_name].reason == REASON_MANUAL

        self.store.remove_exception(rec.normalized_name)
        assert rec.normalized_name not in self.store.load_exceptions()

    def test_duplicate_rejected(self):
        self.store.add_exception(self.name, REASON_MANUAL)
        with pytest.raises(DuplicateExceptionError):
            self.store.add_exception(self.name.lower(), REASON_MANUAL)

    def test_remove_missing_raises(self):
        with pytest.raises(ExceptionNotFoundError):
            self.store.remove_exception("NO SUCH SUPPLIER AT ALL")
```

- [ ] **Step 6: Run tests (DB tests skip)**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: PASS, 3 skipped.

Then run them for real against Neon:
```bash
TEST_DATABASE_URL="$(python -c "import tomllib;print(tomllib.load(open('.streamlit/secrets.toml','rb'))['postgres']['dsn'])")" python -m pytest tests/test_supplier_exceptions.py -q
```
Expected: all pass, 0 skipped.

- [ ] **Step 7: Commit**

```bash
git add src/supplier_exceptions.py tests/test_supplier_exceptions.py
git commit -m "feat: add Neon-backed supplier exceptions store and classifier"
```

---

### Task 4: Seed script

**Files:**
- Create: `scripts/seed_supplier_exceptions.py`

**Interfaces:**
- Consumes: `read_tracker_exceptions` (Task 2), `ExceptionStore` (Task 3).

Doing this now, before the engine work, means the table has real data to develop the next tasks against.

- [ ] **Step 1: Write the script**

Create `scripts/seed_supplier_exceptions.py`:

```python
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
```

- [ ] **Step 2: Dry-run it against the real tracker**

Run:
```bash
python scripts/seed_supplier_exceptions.py "C:\Users\melgh\Downloads\Master Inbound Delivery  Compliance Tracker_April 2025.xlsx" --dry-run
```
Expected: `Found 25 exception supplier(s) in the tracker.` then 25 `would add:` lines.

- [ ] **Step 3: Seed for real**

Run the same command without `--dry-run`.
Expected: `Done. 25 added, 0 already present.` / `Table now holds 25 exception(s).`

- [ ] **Step 4: Verify idempotence**

Run it a second time.
Expected: `Done. 0 added, 25 already present.` If it adds duplicates, the unique index is missing — stop and fix.

- [ ] **Step 5: Commit**

```bash
git add scripts/seed_supplier_exceptions.py
git commit -m "feat: add idempotent seed script for supplier exceptions"
```

---

### Task 5: Engine — Exception Status on Supplier Summary

**Files:**
- Modify: `src/compliance_engine.py`
- Test: `tests/test_supplier_exceptions.py` (append)

**Interfaces:**
- Consumes: `classify_supplier`, `ExceptionRecord` (Task 3).
- Produces: `build_report(sap_df, portal_df, report_year, report_month, exceptions=None, tracker_names=None)` — the two new parameters are keyword, optional, and default to "no exceptions known".

**This is the task where the global constraint bites.** Bill-back and the compliance percentage must come out identical whether or not `exceptions` is passed. Write the regression test FIRST.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_supplier_exceptions.py`:

```python
from src.compliance_engine import build_report


def _sap_row(po, vendor_num, vendor_name, inbound="IBD-1"):
    return {
        "PO Number": po,
        "Normalized PO Number": po,
        "Vendor Number": vendor_num,
        "Vendor Name": vendor_name,
        "Warehouse": "WH1",
        "PO Status": "A",
        "Appointment Date": pd.Timestamp("2026-06-15"),
        "Delivery Date": pd.Timestamp("2026-06-15"),
        "Confirmed PU Date": pd.NaT,
        "Est PU Date": pd.NaT,
        "Inbound Delivery": inbound,
        "Inbound Delivery Status": "A",
    }


def _portal_row(po, supplier, status="Approved"):
    return {
        "PO Number": po,
        "Normalized PO Number": po,
        "Supplier Name": supplier,
        "Upload Date": pd.Timestamp("2026-06-16"),
        "File Status": status,
        "File Name": "doc.pdf",
        "Uploaded By": "someone",
        "Invalid Comment": "",
        "Downloaded By": "",
        "Download Date": pd.NaT,
    }


@pytest.fixture
def scenario():
    """Three suppliers:
      BOTHWELL CHEESE  -- an exception, uploaded nothing
      ACQUA MINERALE   -- on the tracker, uploaded nothing  <- should be chased
      AMERICOLD TACOMA -- not on the tracker, uploaded nothing
      GOOD SUPPLIER    -- on the tracker, uploaded its file
    """
    sap = pd.DataFrame([
        _sap_row("1001", "70001111", "BOTHWELL CHEESE"),
        _sap_row("1002", "70007212", "ACQUA MINERALE SAN BENEDETTO"),
        _sap_row("1003", "70007212", "ACQUA MINERALE SAN BENEDETTO"),
        _sap_row("1004", "70009999", "AMERICOLD TACOMA"),
        _sap_row("1005", "70002222", "GOOD SUPPLIER"),
    ])
    portal = pd.DataFrame([_portal_row("1005", "GOOD SUPPLIER")])
    exceptions = {
        normalize_supplier_name("BOTHWELL CHEESE"): _rec("BOTHWELL CHEESE"),
    }
    tracker = {
        normalize_supplier_name("BOTHWELL CHEESE"),
        normalize_supplier_name("ACQUA MINERALE SAN BENEDETTO"),
        normalize_supplier_name("GOOD SUPPLIER"),
    }
    return sap, portal, exceptions, tracker


class TestSupplierSummaryExceptionColumn:
    def test_column_present_with_the_three_states(self, scenario):
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        summary = sheets["Supplier Summary"].set_index("Vendor Name")

        assert "Exception Status" in sheets["Supplier Summary"].columns
        assert summary.loc["BOTHWELL CHEESE", "Exception Status"] == (
            EXCEPTION_STATUS_EXCEPTION
        )
        assert summary.loc["ACQUA MINERALE SAN BENEDETTO", "Exception Status"] == (
            EXCEPTION_STATUS_EXPECTED
        )
        assert summary.loc["AMERICOLD TACOMA", "Exception Status"] == (
            EXCEPTION_STATUS_NOT_ON_TRACKER
        )

    def test_without_exceptions_column_still_exists(self, scenario):
        # No DB / no exceptions passed: the column reads "Expected to upload"
        # for everyone rather than vanishing, so the sheet's shape is stable.
        sap, portal, _, _ = scenario
        sheets = build_report(sap, portal, 2026, 6)
        col = sheets["Supplier Summary"]["Exception Status"]
        assert set(col) == {EXCEPTION_STATUS_EXPECTED}


class TestBillbackAndComplianceUnchanged:
    """The load-bearing guarantee: exceptions are INFORMATIONAL ONLY."""

    def test_billback_identical_with_and_without_exceptions(self, scenario):
        sap, portal, exceptions, tracker = scenario
        without = build_report(sap, portal, 2026, 6)
        with_exc = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )

        bb_without = {k: v for k, v in without.items() if k.startswith("BB-")}
        bb_with = {k: v for k, v in with_exc.items() if k.startswith("BB-")}

        assert set(bb_without) == set(bb_with)
        # Bothwell is an exception but is STILL billed -- by design, for now.
        assert any("BOTHWELL" in k.upper() for k in bb_with)
        for name in bb_without:
            pd.testing.assert_frame_equal(bb_without[name], bb_with[name])

    def test_compliance_percentage_identical(self, scenario):
        sap, portal, exceptions, tracker = scenario
        without = build_report(sap, portal, 2026, 6)
        with_exc = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        pd.testing.assert_frame_equal(
            without["Monthly Summary"], with_exc["Monthly Summary"]
        )


class TestEmptyFrames:
    def test_no_sap_rows_does_not_crash(self, scenario):
        """pandas 3.0: empty .map/.apply blow up. Guard the exceptions path too."""
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap.iloc[0:0], portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        assert sheets["Supplier Summary"].empty
        assert sheets["Should Have Uploaded"].empty
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: FAIL — `TypeError: build_report() got an unexpected keyword argument 'exceptions'`

- [ ] **Step 3: Add the exceptions import to the engine**

In `src/compliance_engine.py`, extend the existing `from .config import (...)` block with:

```python
    EXCEPTION_STATUS_EXCEPTION,
    EXCEPTION_STATUS_EXPECTED,
```

and add below the existing `from .normalizer import has_value`:

```python
from .supplier_exceptions import ExceptionRecord, classify_supplier
```

- [ ] **Step 4: Change the signature and annotate the rows**

Replace the `build_report` signature (`compliance_engine.py:68-73`) with:

```python
def build_report(
    sap_df: pd.DataFrame,
    portal_df: pd.DataFrame,
    report_year: int,
    report_month: int,
    exceptions: dict[str, "ExceptionRecord"] | None = None,
    tracker_names: set[str] | None = None,
) -> dict[str, pd.DataFrame]:
    """Return a dict {sheet_name: dataframe} for every required sheet.

    `exceptions` maps normalized supplier name -> ExceptionRecord; `tracker_names`
    is every supplier the tracker knows about. Both default to empty, in which
    case every supplier is labelled "Expected to upload" and the report is
    identical to one built before this feature existed.

    Exceptions are INFORMATIONAL. They deliberately do not affect bill-back or
    the compliance percentage.
    """
    exceptions = exceptions or {}
    tracker_names = tracker_names or set()
```

Then, immediately after the line that sets `sap_valid["Has Inbound"]` (`compliance_engine.py:100`), add:

```python
    # Exception Status is annotated per SAP row (not just in the summary rollup)
    # so a future change can act on it without re-plumbing the engine.
    sap_valid["Exception Status"] = [
        classify_supplier(name, number, exceptions, tracker_names)
        for name, number in zip(
            sap_valid["Vendor Name"], sap_valid["Vendor Number"], strict=True
        )
    ]
```

A list comprehension, not `.apply(axis=1)` — it is empty-frame safe by construction, which `apply` is not under pandas 3.0.

- [ ] **Step 5: Add the column to the Supplier Summary**

In `_supplier_summary` (`compliance_engine.py:537`), the groupby is over `["Vendor Number", "Vendor Name"]`, so every row in a group shares one Exception Status. Inside the loop, after `processing_n = ...`, add:

```python
        status = (
            g["Exception Status"].iloc[0]
            if "Exception Status" in g.columns and len(g)
            else EXCEPTION_STATUS_EXPECTED
        )
```

and add to the appended dict, immediately after `"Vendor Name": vendor_name,`:

```python
            "Exception Status": status,
```

- [ ] **Step 6: Run tests**

Run: `python -m pytest tests/test_supplier_exceptions.py -q -k "SupplierSummary or Unchanged"`
Expected: the `TestSupplierSummaryExceptionColumn` and `TestBillbackAndComplianceUnchanged` tests PASS. `TestEmptyFrames` still fails — `Should Have Uploaded` doesn't exist yet. That's Task 6.

- [ ] **Step 7: Confirm the whole existing suite is untouched**

Run: `python -m pytest tests/ -q`
Expected: every pre-existing test still passes. If any existing test broke, the "no behaviour change" constraint has been violated — fix the engine, do NOT edit the old test.

- [ ] **Step 8: Commit**

```bash
git add src/compliance_engine.py tests/test_supplier_exceptions.py
git commit -m "feat: add Exception Status column to the Supplier Summary sheet"
```

---

### Task 6: The "Should Have Uploaded" sheet

**Files:**
- Modify: `src/compliance_engine.py`
- Test: `tests/test_supplier_exceptions.py` (append)

**Interfaces:**
- Produces: a `"Should Have Uploaded"` key in `build_report`'s returned dict.

A supplier lands on this sheet when **all** hold: at least one SAP PO with an inbound delivery this month; **zero** portal uploads of any kind (no valid *and* no Invalid — a rejected upload still means they know the process exists); and `Exception Status != "Exception"`.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_supplier_exceptions.py`:

```python
class TestShouldHaveUploaded:
    def test_lists_only_non_exception_zero_upload_suppliers(self, scenario):
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        names = set(sheets["Should Have Uploaded"]["Vendor Name"])

        assert "ACQUA MINERALE SAN BENEDETTO" in names  # expected, uploaded nothing
        assert "AMERICOLD TACOMA" in names              # not on tracker, still chased
        assert "BOTHWELL CHEESE" not in names           # an exception
        assert "GOOD SUPPLIER" not in names             # it uploaded

    def test_sorted_by_most_pos_first(self, scenario):
        sap, portal, exceptions, tracker = scenario
        sheets = build_report(
            sap, portal, 2026, 6, exceptions=exceptions, tracker_names=tracker
        )
        counts = sheets["Should Have Uploaded"]["Inbound POs Expected"].tolist()
        assert counts == sorted(counts, reverse=True)
        # Acqua has 2 POs, Americold 1.
        top = sheets["Should Have Uploaded"].iloc[0]
        assert top["Vendor Name"] == "ACQUA MINERALE SAN BENEDETTO"
        assert top["Inbound POs Expected"] == 2
        assert top["Portal Uploads"] == 0
        assert top["Bill-Back Total"] == 400  # 2 POs x $200

    def test_a_supplier_with_an_invalid_upload_is_excluded(self):
        """An Invalid upload still means the supplier knows the process exists.
        They belong in bill-back, not on the 'never even tried' list."""
        sap = pd.DataFrame([_sap_row("2001", "70003333", "TRIED AND FAILED")])
        portal = pd.DataFrame([_portal_row("2001", "TRIED AND FAILED", status="Invalid")])
        sheets = build_report(sap, portal, 2026, 6)
        assert "TRIED AND FAILED" not in set(sheets["Should Have Uploaded"]["Vendor Name"])

    def test_a_supplier_with_no_inbound_delivery_is_excluded(self):
        """No inbound delivery means there was nothing to document yet."""
        sap = pd.DataFrame([_sap_row("3001", "70004444", "NOT SHIPPED YET", inbound="")])
        sheets = build_report(sap, pd.DataFrame(columns=list(_portal_row("x", "y"))), 2026, 6)
        assert "NOT SHIPPED YET" not in set(sheets["Should Have Uploaded"]["Vendor Name"])
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_supplier_exceptions.py -q -k ShouldHaveUploaded`
Expected: FAIL — `KeyError: 'Should Have Uploaded'`

- [ ] **Step 3: Write the builder**

Add to `src/compliance_engine.py`, next to `_supplier_summary` in the "Group rollups" section:

```python
def _should_have_uploaded(sap_unique: pd.DataFrame) -> pd.DataFrame:
    """Suppliers who uploaded NOTHING despite being expected to.

    Stricter, and more damning, than "missing some": a supplier who uploaded 9 of
    10 POs has a working process with a gap; one who uploaded 0 of 10 does not
    know the process exists. Partial cases are already covered by the bill-back
    tabs, so this sheet earns its place only by isolating total failures.

    An Invalid (rejected) upload counts as "they tried" and keeps a supplier OFF
    this sheet.
    """
    columns = [
        "Vendor Number", "Vendor Name", "Exception Status",
        "Inbound POs Expected", "Portal Uploads", "Bill-Back Total",
    ]
    if sap_unique.empty:
        return pd.DataFrame(columns=columns)

    rows = []
    for (vendor_num, vendor_name), g in sap_unique.groupby(
        ["Vendor Number", "Vendor Name"], dropna=False
    ):
        status = g["Exception Status"].iloc[0]
        if status == EXCEPTION_STATUS_EXCEPTION:
            continue

        uploads = int(
            g[g["Portal Match"] | g["Portal Invalid Match"]][
                "Normalized PO Number"
            ].nunique()
        )
        if uploads:
            continue

        expected = int(g[g["Has Inbound"]]["Normalized PO Number"].nunique())
        if not expected:
            continue

        rows.append({
            "Vendor Number": vendor_num,
            "Vendor Name": vendor_name,
            "Exception Status": status,
            "Inbound POs Expected": expected,
            "Portal Uploads": 0,
            "Bill-Back Total": expected * BILLBACK_FEE_PER_OCCURRENCE,
        })

    if not rows:
        return pd.DataFrame(columns=columns)
    return (
        pd.DataFrame(rows)
        .sort_values("Inbound POs Expected", ascending=False, kind="stable")
        .reset_index(drop=True)
    )
```

- [ ] **Step 4: Register the sheet**

In the `sheets = {...}` dict in `build_report` (`compliance_engine.py:271-285`), add immediately after the `"Supplier Summary"` entry:

```python
        "Should Have Uploaded": _should_have_uploaded(sap_unique),
```

It goes right after Supplier Summary so the chase-list sits next to the rollup it derives from, and before the `BB-` tabs that `sheets.update(...)` appends.

- [ ] **Step 5: Run tests**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: PASS, DB tests skipped. `TestEmptyFrames` now passes too.

- [ ] **Step 6: Run the whole suite**

Run: `python -m pytest tests/ -q`
Expected: all pre-existing tests still pass.

- [ ] **Step 7: Commit**

```bash
git add src/compliance_engine.py tests/test_supplier_exceptions.py
git commit -m "feat: add Should Have Uploaded sheet for zero-upload suppliers"
```

---

### Task 7: Dashboard wiring

**Files:**
- Modify: `pages/1_Supplier_Compliance_Dashboard.py`
- Create: `src/supplier_exceptions_ui.py`
- Test: `tests/test_supplier_exceptions.py` (append)

**Interfaces:**
- Consumes: `ExceptionStore` (Task 3).
- Produces: `load_exceptions_or_empty() -> tuple[dict[str, ExceptionRecord], set[str], str | None]` — returns `(exceptions, tracker_names, error_message)`. `error_message` is `None` on success.

**`tracker_names` on the dashboard:** the app has no tracker workbook at runtime — only the DB. So the dashboard passes `tracker_names=set()`, which collapses `Not on tracker` into `Expected to upload`. The three-state distinction is available to the engine and the tests, but the live dashboard shows two states until someone loads a tracker. This is intentional YAGNI: adding a tracker uploader is a separate, easy change, and the user asked for the column, not a second uploader.

- [ ] **Step 1: Write the UI helper**

Create `src/supplier_exceptions_ui.py`:

```python
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
```

- [ ] **Step 2: Wire the page**

In `pages/1_Supplier_Compliance_Dashboard.py`, add to the imports:

```python
from src.supplier_exceptions_ui import load_exceptions_or_empty, render_exception_manager
```

Replace the `build_report` call (`pages/1_Supplier_Compliance_Dashboard.py:88-89`) with:

```python
    exceptions, tracker_names, exceptions_error = load_exceptions_or_empty()
    if exceptions_error:
        st.info(exceptions_error)

    with st.spinner("Applying compliance rules..."):
        sheets = build_report(
            sap_df, portal_df, sel_year, sel_month,
            exceptions=exceptions, tracker_names=tracker_names,
        )
```

Then, immediately after the existing bill-back block (the `if billback_tabs: ... else: ...` at lines 103-115) and before the "Writing Excel workbook" spinner, add:

```python
    chase = sheets["Should Have Uploaded"]
    st.subheader("Should Have Uploaded — Nothing Received")
    if chase.empty:
        st.caption("Every supplier expected to upload submitted at least one file.")
    else:
        st.caption(
            f"**{len(chase)}** supplier(s) uploaded **nothing at all** this month "
            "despite having inbound deliveries, and are not on the exceptions list."
        )
        st.dataframe(chase, use_container_width=True, hide_index=True)
```

Finally, at the very end of the file (outside the `if st.button(...)` block, so it is always reachable — a Streamlit button is True only on the run that clicks it):

```python
render_exception_manager()
```

- [ ] **Step 3: Write the smoke test**

Append to `tests/test_supplier_exceptions.py`:

```python
from streamlit.testing.v1 import AppTest


class TestDashboardBoots:
    def test_page_1_still_renders(self):
        """The exceptions panel must not crash the page when no DB is configured."""
        at = AppTest.from_file("app.py", default_timeout=30).run()
        assert not at.exception
```

- [ ] **Step 4: Run the tests**

Run: `python -m pytest tests/test_supplier_exceptions.py -q`
Expected: PASS, DB tests skipped.

- [ ] **Step 5: Verify in the real app**

Run: `python -m streamlit run app.py`

Then in the browser: upload SAP `C:\Users\melgh\Downloads\export (18).xlsx` and Portal `C:\Users\melgh\Downloads\inbound-delivery-file-upload-audit.xlsx`, set **Report Month = June 2026** (July is empty and hits a known pandas-3 crash in `portal_importer.load_portal`), and click Generate.

Confirm:
- Supplier Summary has an `Exception Status` column, and the 5 seeded suppliers that appear in this export (Agropur, Bothwell Cheese, Caffe Mauro, Mangiatorella, Serfunghi) read `Exception`.
- The "Should Have Uploaded" section renders and excludes those 5.
- The bill-back line still reads **59 suppliers billed — total $18,800**, exactly as before this feature. If that number moved, the informational-only constraint is broken.
- "Manage Supplier Exceptions" expands and lists 25 suppliers.

- [ ] **Step 6: Commit**

```bash
git add pages/1_Supplier_Compliance_Dashboard.py src/supplier_exceptions_ui.py tests/test_supplier_exceptions.py
git commit -m "feat: surface supplier exceptions on the compliance dashboard"
```

---

### Task 8: Docs + final verification

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Run the full suite, including the DB tests**

```bash
TEST_DATABASE_URL="$(python -c "import tomllib;print(tomllib.load(open('.streamlit/secrets.toml','rb'))['postgres']['dsn'])")" python -m pytest tests/ -q
```
Expected: everything passes, 0 skipped.

- [ ] **Step 2: Document the feature**

Add a `## Supplier Exceptions` section to `README.md` covering: what an exception is; that the list is seeded from the tracker by `scripts/seed_supplier_exceptions.py` but the **DB is the source of truth afterwards** (editing the Excel does not update the app); that the status is **informational and does not change bill-back or the compliance percentage**; and that the `[postgres] dsn` secret must be set in Streamlit Cloud for the column to populate on the live app.

- [ ] **Step 3: Commit and push**

```bash
git add README.md
git commit -m "docs: document the supplier exceptions feature"
git push -u origin supplier-exceptions
```

---

## Deployment note

The live app reads `st.secrets["postgres"]["dsn"]` from **Streamlit Cloud → Settings → Secrets**. Per the project memory, this secret may still be unset in Cloud — in which case the column-variants feature shows Standard-only and this feature will show the "No database configured" info message. Setting it once fixes both. The app never crashes either way.
