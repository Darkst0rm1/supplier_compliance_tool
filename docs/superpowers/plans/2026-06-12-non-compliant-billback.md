# Non-Compliant Bill-Back Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add one Excel tab per non-compliant supplier to the compliance workbook, listing PO numbers whose inbound documentation was never uploaded, each charged $200.

**Architecture:** A new pure helper in `compliance_engine.py` filters the existing `missing` bucket down to never-uploaded POs, groups them by supplier, and returns an ordered `{sheet_name: DataFrame}` dict merged into `build_report`'s output. `report_generator.generate_workbook` already renders any sheet dict, so no writer changes are needed. Fee and reason live in `config.py`.

**Tech Stack:** Python 3.14, pandas, openpyxl, Streamlit; pytest for unit tests (new dev dependency).

**Spec:** `docs/superpowers/specs/2026-06-12-non-compliant-billback-design.md`

**Python interpreter:** `C:\Users\melgh\AppData\Local\Python\pythoncore-3.14-64\python.exe` (not on PATH). All commands below use `python` as shorthand — substitute the full path. Run from repo root `C:\Users\melgh\Documents\GitHub\supplier_compliance_tool`.

---

## File Structure

- `src/config.py` — **modify**: add `BILLBACK_FEE_PER_OCCURRENCE` and `BILLBACK_REASON`. (Already holds the `Received` status fix in the working tree — commit that first.)
- `src/compliance_engine.py` — **modify**: add `_billback_sheet_name`, `_billback_supplier_tab`, `_billback_sheets`; merge bill-back sheets into `build_report`'s return dict.
- `pages/1_Supplier_Compliance_Dashboard.py` — **modify**: surface a small "suppliers billed / total $" caption after report generation (display only).
- `requirements-dev.txt` — **create**: pinned pytest for the unit tests.
- `tests/test_billback.py` — **create**: unit tests for the three new helpers.
- `tests/__init__.py` — **create**: empty, makes `tests` importable.

---

## Task 1: Dev tooling + config constants

**Files:**
- Create: `requirements-dev.txt`
- Modify: `src/config.py`

- [ ] **Step 1: Commit the pending `Received` status fix first**

The working tree already contains the `Received` → compliant change in `config.py` from the prior session. Commit it on its own so the bill-back work starts clean.

```bash
git add src/config.py
git commit -m "fix: treat portal 'Received' status as compliant like Approved"
```

- [ ] **Step 2: Create the dev requirements file**

Create `requirements-dev.txt`:

```
-r requirements.txt
pytest>=8.0
```

- [ ] **Step 3: Install pytest**

Run: `python -m pip install -r requirements-dev.txt`
Expected: pytest installs successfully (`Successfully installed pytest-...`).

- [ ] **Step 4: Verify pytest runs**

Run: `python -m pytest --version`
Expected: prints `pytest 8.x.x`.

- [ ] **Step 5: Add bill-back constants to config**

In `src/config.py`, after the `PORTAL_PENDING_STATUSES` block, add:

```python
# Bill-back: suppliers are charged a flat fee for every inbound PO whose
# documentation was never uploaded to the portal (a "Missing Inbound Document").
# Uploaded-but-Invalid POs are NOT billed -- the supplier attempted.
BILLBACK_FEE_PER_OCCURRENCE = 200          # USD per missing inbound document
BILLBACK_REASON = "Missing Inbound Document"
```

- [ ] **Step 6: Commit**

```bash
git add requirements-dev.txt src/config.py
git commit -m "chore: add pytest dev dep and bill-back fee constants"
```

---

## Task 2: Sheet-name sanitizer

**Files:**
- Modify: `src/compliance_engine.py`
- Create: `tests/__init__.py`, `tests/test_billback.py`

Excel sheet names cap at 31 chars and forbid `: \ / ? * [ ]`. Two different vendors can sanitize to the same name, so collisions get a numeric suffix.

- [ ] **Step 1: Create the test package marker**

Create `tests/__init__.py` as an empty file (one blank line is fine).

- [ ] **Step 2: Write the failing tests**

Create `tests/test_billback.py`:

```python
from src.compliance_engine import _billback_sheet_name


def test_sheet_name_prefixes_and_keeps_simple_name():
    used = set()
    assert _billback_sheet_name("BOB'S RED MILL", "1001", used) == "BB-BOB'S RED MILL"


def test_sheet_name_strips_illegal_chars():
    used = set()
    name = _billback_sheet_name("A/B:C*D[E]F?G\\H", "1", used)
    for ch in r":\/?*[]":
        assert ch not in name
    assert name.startswith("BB-")


def test_sheet_name_truncates_to_31_chars():
    used = set()
    name = _billback_sheet_name("X" * 60, "1", used)
    assert len(name) <= 31
    assert name.startswith("BB-")


def test_sheet_name_dedupes_collisions():
    used = set()
    first = _billback_sheet_name("SAME NAME", "1", used)
    second = _billback_sheet_name("SAME NAME", "2", used)
    assert first != second
    assert first in used and second in used
    assert len(second) <= 31


def test_sheet_name_falls_back_to_number_then_unknown():
    used = set()
    assert _billback_sheet_name("", "9999", used) == "BB-9999"
    assert _billback_sheet_name("", "", used) == "BB-Unknown Supplier"
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `python -m pytest tests/test_billback.py -v`
Expected: FAIL — `ImportError: cannot import name '_billback_sheet_name'`.

- [ ] **Step 4: Implement the sanitizer**

In `src/compliance_engine.py`, add near the other private helpers (e.g. after `_processing_review_status`):

```python
_ILLEGAL_SHEET_CHARS = set(r":\/?*[]")


def _billback_sheet_name(vendor_name: str, vendor_number: str, used: set) -> str:
    """Return a unique, Excel-legal bill-back sheet name (<=31 chars).

    Prefixes with 'BB-' so all bill-back tabs group together. Falls back to the
    vendor number, then 'Unknown Supplier', when the name is blank. Collisions
    against names already in `used` get a numeric suffix. Mutates `used`.
    """
    base = (vendor_name or "").strip()
    if not base:
        base = (vendor_number or "").strip()
    if not base:
        base = "Unknown Supplier"
    base = "".join(" " if c in _ILLEGAL_SHEET_CHARS else c for c in base)
    base = " ".join(base.split())  # collapse runs of whitespace

    name = ("BB-" + base)[:31]
    if name in used:
        i = 2
        while True:
            suffix = f"-{i}"
            name = ("BB-" + base)[: 31 - len(suffix)] + suffix
            if name not in used:
                break
            i += 1
    used.add(name)
    return name
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `python -m pytest tests/test_billback.py -v`
Expected: all 5 tests PASS.

- [ ] **Step 6: Commit**

```bash
git add src/compliance_engine.py tests/__init__.py tests/test_billback.py
git commit -m "feat: add bill-back sheet-name sanitizer"
```

---

## Task 3: Per-supplier tab builder

**Files:**
- Modify: `src/compliance_engine.py`, `tests/test_billback.py`

Builds one supplier's tab: the billable columns plus a `Charge (USD)` of 200 per row, then a TOTAL row.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_billback.py`:

```python
import pandas as pd

from src.compliance_engine import _billback_supplier_tab
from src.config import BILLBACK_FEE_PER_OCCURRENCE, BILLBACK_REASON


def _sample_missing_rows():
    return pd.DataFrame(
        {
            "PO Number": ["1001", "1002"],
            "Warehouse": ["W1", "W1"],
            "PO Status": ["B", "B"],
            "Appointment Date": ["2026-06-01", "2026-06-02"],
            "Delivery Date": ["2026-06-03", "2026-06-04"],
            "Inbound Delivery": ["IB1", "IB2"],
        }
    )


def test_supplier_tab_has_expected_columns():
    tab = _billback_supplier_tab(_sample_missing_rows())
    assert list(tab.columns) == [
        "PO Number", "Warehouse", "PO Status", "Appointment Date",
        "Delivery Date", "Inbound Delivery", "Charge Reason", "Charge (USD)",
    ]


def test_supplier_tab_charges_fee_per_po():
    tab = _billback_supplier_tab(_sample_missing_rows())
    # First 2 rows are POs, last row is the TOTAL row.
    po_rows = tab.iloc[:-1]
    assert (po_rows["Charge (USD)"] == BILLBACK_FEE_PER_OCCURRENCE).all()
    assert (po_rows["Charge Reason"] == BILLBACK_REASON).all()


def test_supplier_tab_total_row():
    tab = _billback_supplier_tab(_sample_missing_rows())
    total = tab.iloc[-1]
    assert "2 occurrences" in str(total["PO Number"])
    assert total["Charge (USD)"] == 2 * BILLBACK_FEE_PER_OCCURRENCE
```

- [ ] **Step 2: Run the new tests to verify they fail**

Run: `python -m pytest tests/test_billback.py -k supplier_tab -v`
Expected: FAIL — `ImportError: cannot import name '_billback_supplier_tab'`.

- [ ] **Step 3: Implement the tab builder**

In `src/compliance_engine.py`, add the import at the top alongside the other `from .config import (...)` names: add `BILLBACK_FEE_PER_OCCURRENCE` and `BILLBACK_REASON`. Then add the helper:

```python
def _billback_supplier_tab(rows: pd.DataFrame) -> pd.DataFrame:
    """One supplier's bill-back tab: billable POs + a TOTAL row."""
    tab = pd.DataFrame(
        {
            "PO Number": rows["PO Number"].astype(str).values,
            "Warehouse": rows["Warehouse"].values,
            "PO Status": rows["PO Status"].values,
            "Appointment Date": rows["Appointment Date"].values,
            "Delivery Date": rows["Delivery Date"].values,
            "Inbound Delivery": rows["Inbound Delivery"].values,
            "Charge Reason": BILLBACK_REASON,
            "Charge (USD)": BILLBACK_FEE_PER_OCCURRENCE,
        }
    )
    n = len(tab)
    total = {c: "" for c in tab.columns}
    total["PO Number"] = f"TOTAL — {n} occurrences"
    total["Charge (USD)"] = n * BILLBACK_FEE_PER_OCCURRENCE
    return pd.concat([tab, pd.DataFrame([total])], ignore_index=True)
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_billback.py -v`
Expected: all tests (Task 2 + Task 3) PASS.

- [ ] **Step 5: Commit**

```bash
git add src/compliance_engine.py tests/test_billback.py
git commit -m "feat: add per-supplier bill-back tab builder"
```

---

## Task 4: Bill-back orchestrator + wire into build_report

**Files:**
- Modify: `src/compliance_engine.py`, `tests/test_billback.py`

Filters the `missing` bucket to never-uploaded rows, groups by supplier (Vendor Number, falling back to name), orders suppliers by occurrence count descending, and returns `{sheet_name: tab}`.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_billback.py`:

```python
from src.compliance_engine import _billback_sheets


def _missing_bucket():
    # Two never-uploaded POs for vendor 1001, one for vendor 2002,
    # plus one Invalid-upload PO that must be excluded from billing.
    return pd.DataFrame(
        {
            "PO Number": ["1001", "1002", "2001", "9001"],
            "Vendor Number": ["1001", "1001", "2002", "1001"],
            "Vendor Name": ["BOB'S RED MILL", "BOB'S RED MILL",
                            "HP HOOD LLC", "BOB'S RED MILL"],
            "Warehouse": ["W1", "W1", "W2", "W1"],
            "PO Status": ["B", "B", "B", "B"],
            "Appointment Date": ["2026-06-01"] * 4,
            "Delivery Date": ["2026-06-03"] * 4,
            "Inbound Delivery": ["IB1", "IB2", "IB3", "IB9"],
            "Portal Invalid Match": [False, False, False, True],
        }
    )


def test_billsheets_excludes_invalid_uploads():
    sheets = _billback_sheets(_missing_bucket())
    # Vendor 1001 should bill only its 2 never-uploaded POs, not the Invalid one.
    bob_name = next(n for n in sheets if "BOB" in n)
    bob_tab = sheets[bob_name]
    po_rows = bob_tab.iloc[:-1]
    assert set(po_rows["PO Number"]) == {"1001", "1002"}
    assert "9001" not in set(po_rows["PO Number"])


def test_billsheets_one_tab_per_supplier():
    sheets = _billback_sheets(_missing_bucket())
    assert len(sheets) == 2  # BOB'S RED MILL + HP HOOD LLC


def test_billsheets_ordered_by_occurrences_desc():
    sheets = _billback_sheets(_missing_bucket())
    first_name = next(iter(sheets))  # dict preserves insertion order
    assert "BOB" in first_name  # 2 occurrences ranks above HP HOOD's 1


def test_billsheets_empty_when_no_billable():
    empty = _missing_bucket().iloc[0:0]
    assert _billback_sheets(empty) == {}


def test_billsheets_empty_when_all_invalid():
    df = _missing_bucket().copy()
    df["Portal Invalid Match"] = True
    assert _billback_sheets(df) == {}
```

- [ ] **Step 2: Run the new tests to verify they fail**

Run: `python -m pytest tests/test_billback.py -k billsheets -v`
Expected: FAIL — `ImportError: cannot import name '_billback_sheets'`.

- [ ] **Step 3: Implement the orchestrator**

In `src/compliance_engine.py`, add:

```python
def _billback_sheets(missing: pd.DataFrame) -> dict:
    """Build {sheet_name: tab} for every supplier with never-uploaded POs.

    Only rows whose portal file was never submitted are billed; rows with an
    Invalid (rejected) upload are excluded. Suppliers are ordered by occurrence
    count descending so the biggest offenders' tabs come first.
    """
    if missing is None or missing.empty:
        return {}

    billable = missing
    if "Portal Invalid Match" in billable.columns:
        billable = billable[~billable["Portal Invalid Match"].fillna(False)]
    billable = billable.copy()
    if billable.empty:
        return {}

    vnum = billable["Vendor Number"].fillna("").astype(str).str.strip()
    vname = billable["Vendor Name"].fillna("").astype(str).str.strip()
    key = vnum.where(vnum != "", vname)
    key = key.where(key != "", "Unknown Supplier")
    billable = billable.assign(__vkey=key.values, __vname=vname.values)

    order = (
        billable.groupby("__vkey").size().sort_values(ascending=False).index.tolist()
    )

    used: set = set()
    sheets: dict = {}
    for vkey in order:
        rows = billable[billable["__vkey"] == vkey]
        display_name = next((n for n in rows["__vname"] if n), str(vkey))
        sheet_name = _billback_sheet_name(display_name, str(vkey), used)
        sheets[sheet_name] = _billback_supplier_tab(rows)
    return sheets
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest tests/test_billback.py -v`
Expected: all tests PASS.

- [ ] **Step 5: Merge bill-back sheets into `build_report`**

In `src/compliance_engine.py`, the `build_report` function currently ends with `return { ... }`. Change that final block from:

```python
    return {
        "Monthly Summary": summary,
        ...
        "Warehouse Summary": _warehouse_summary(sap_unique),
    }
```

to assign the dict then append bill-back tabs:

```python
    sheets = {
        "Monthly Summary": summary,
        "Portal Export Data": _portal_sheet(portal),
        "SAP Export Data": _sap_sheet(sap),
        "SAP Inbound Matched With Portal File": _matched_columns(matched),
        "SAP Inbound Missing Portal File": _missing_columns(missing),
        "Pending TOL Review": _pending_columns(pending_unique),
        "Portal File But No SAP Inbound": _portal_no_inbound_columns(portal_no_inbound),
        "Portal PO Not Found In SAP": _not_in_sap_columns(not_in_sap),
        "No Inbound Yet": _no_inbound_yet_columns(no_inbound_yet),
        "Closed POs Review": _review_columns(closed),
        "Processing POs Review": _review_columns(processing),
        "Supplier Summary": _supplier_summary(sap_unique),
        "Warehouse Summary": _warehouse_summary(sap_unique),
    }
    sheets.update(_billback_sheets(missing))
    return sheets
```

Note: `_billback_sheets` receives the full `missing` dataframe (which still has the
`Vendor Number`, `Vendor Name`, and `Portal Invalid Match` columns) — call it with
`missing`, not the column-narrowed `_missing_columns(missing)`.

- [ ] **Step 6: Run the full test suite to verify nothing broke**

Run: `python -m pytest tests/ -v`
Expected: all tests PASS.

- [ ] **Step 7: Commit**

```bash
git add src/compliance_engine.py tests/test_billback.py
git commit -m "feat: build per-supplier bill-back tabs into compliance workbook"
```

---

## Task 5: Dashboard caption + end-to-end verification

**Files:**
- Modify: `pages/1_Supplier_Compliance_Dashboard.py`

Surface a small read-out of how many suppliers were billed and the total, using the bill-back tabs already present in `sheets`.

- [ ] **Step 1: Add the bill-back read-out after the summary metrics**

In `pages/1_Supplier_Compliance_Dashboard.py`, after the `m1..m4` metric block (the `m4.metric("Compliance %", ...)` line) and before the `with st.spinner("Writing Excel workbook..."):` block, insert:

```python
    billback_tabs = {k: v for k, v in sheets.items() if k.startswith("BB-")}
    if billback_tabs:
        total_charge = sum(
            int(tab.iloc[-1]["Charge (USD)"]) for tab in billback_tabs.values()
        )
        st.subheader("Non-Compliant Bill-Back")
        st.caption(
            f"{len(billback_tabs)} supplier(s) billed for missing inbound "
            f"documents — total **${total_charge:,}**. One tab per supplier is "
            "included in the Excel download (sheets prefixed `BB-`)."
        )
    else:
        st.caption("No bill-back: every inbound PO had its document uploaded.")
```

- [ ] **Step 2: Launch the app and verify manually**

Run: `python -m streamlit run app.py`
Then in the browser, on the **Supplier Compliance Dashboard** page:
1. Upload a SAP export (`.xlsx`) and the portal export
   `C:\Users\melgh\Downloads\inbound-delivery-file-upload-audit.xlsx`.
2. Pick **June 2026**, click **Generate Compliance Report**.
3. Expected: a "Non-Compliant Bill-Back" caption shows the supplier count and total.
4. Download the Excel; confirm there is one `BB-<supplier>` tab per non-compliant
   supplier, each ending in a TOTAL row, and that no supplier with only an Invalid
   upload appears.

(If no SAP export is on hand, the unit tests in `tests/test_billback.py` already
cover the billing logic; the manual run only validates the UI wiring and Excel output.)

- [ ] **Step 3: Commit**

```bash
git add pages/1_Supplier_Compliance_Dashboard.py
git commit -m "feat: show bill-back summary on compliance dashboard"
```

---

## Task 6: Wrap up

- [ ] **Step 1: Run the full suite one last time**

Run: `python -m pytest tests/ -v`
Expected: all tests PASS.

- [ ] **Step 2: Confirm the branch is clean**

Run: `git status`
Expected: working tree clean on `feature/non-compliant-billback`.

- [ ] **Step 3: Report the diff summary to the user** and ask whether to merge to `main` or open a PR. Do not merge without the user's go-ahead.
```
