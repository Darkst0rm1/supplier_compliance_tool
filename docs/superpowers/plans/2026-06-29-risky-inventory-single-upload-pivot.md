# Risky Inventory — Single 0–180 Upload + Real PivotTable — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the two-file Risky Inventory flow with a single 0–180 upload split by `MRP Last Sell Date`, output a workbook whose Summary is a real interactive Excel PivotTable.

**Architecture:** Pure-Python bucketing on the loaded detail; output built by filling a **committed template workbook** (derived once from the golden file, with a `Bucket` page field added and all supplier data scrubbed) and setting the pivot cache to refresh on open. openpyxl can't create pivots but preserves an existing one through load→save — validated.

**Tech Stack:** Python 3.14, openpyxl, Streamlit, pandas (display only), pytest.

## Global Constraints

- Run tests with the project interpreter: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -m pytest tests/ -q`.
- No new dependencies — openpyxl is already used.
- Output sheet names: **`Detail`** (data) and **`Summary`** (pivot). `Bucket` is the **last** column of `Detail` (column U / 21).
- Bucketing (cutoff = run date + 90 days): `MRP Last Sell Date` **≤ cutoff → `0-90 Day`**; **> cutoff → `91-180 Day`**; **blank/missing → `No Last Sell Date`**. Boundary inclusive on the 0–90 side.
- The committed template asset MUST contain **no real supplier data** (cache `sharedItems` and records scrubbed to empty; rebuilt by Excel on open via `refreshOnLoad`).
- All 8 existing pages must still load.
- Existing detail layout: 20-column `DETAIL_HEADERS` on `Sheet1`; `load_detail` already validates headers and captures per-column number formats / widths / header styles.

---

### Task 1: Bucketing logic + cutoff (engine)

**Files:**
- Modify: `src/risky_inventory_engine.py`
- Test: `tests/test_risky_inventory.py`

**Interfaces:**
- Consumes: existing `DetailTable`, `load_detail`, `DETAIL_HEADERS`.
- Produces:
  - `compute_cutoff(run_date: date) -> date`
  - `bucket_for(last_sell, cutoff: date) -> str`
  - `assign_buckets(detail: DetailTable, cutoff: date) -> tuple[DetailTable, dict[str, int]]`
    (returns a new DetailTable whose `headers == DETAIL_HEADERS + ["Bucket"]` and whose
    rows each have the bucket appended, plus a counts dict keyed by the three bucket labels)
  - Constants: `MRP_LAST_SELL_COL`, `BUCKET_COL`, `BUCKET_0_90`, `BUCKET_91_180`, `BUCKET_NONE`, `CUTOFF_DAYS`.

- [ ] **Step 1: Write failing tests**

Add to `tests/test_risky_inventory.py` (the `_make_row` / `_make_xlsx` / `load_detail` helpers already exist there):

```python
from datetime import date, datetime
from src.risky_inventory_engine import (
    BUCKET_0_90, BUCKET_91_180, BUCKET_NONE,
    assign_buckets, bucket_for, compute_cutoff,
)


def test_compute_cutoff_is_run_date_plus_90():
    assert compute_cutoff(date(2026, 6, 24)) == date(2026, 9, 22)


def test_bucket_for_boundaries():
    cutoff = date(2026, 9, 22)
    assert bucket_for(datetime(2026, 9, 22), cutoff) == BUCKET_0_90   # inclusive
    assert bucket_for(datetime(2026, 9, 21), cutoff) == BUCKET_0_90
    assert bucket_for(datetime(2026, 9, 23), cutoff) == BUCKET_91_180
    assert bucket_for(None, cutoff) == BUCKET_NONE
    assert bucket_for("", cutoff) == BUCKET_NONE


def test_assign_buckets_appends_column_and_counts():
    rows = [
        _make_row(Material="A", **{"MRP Last Sell Date": datetime(2026, 8, 1)}),   # 0-90
        _make_row(Material="B", **{"MRP Last Sell Date": datetime(2026, 12, 1)}),  # 91-180
        _make_row(Material="C", **{"MRP Last Sell Date": None}),                   # none
    ]
    detail = load_detail(_make_xlsx(rows))
    bucketed, counts = assign_buckets(detail, compute_cutoff(date(2026, 6, 24)))
    assert bucketed.headers[-1] == "Bucket"
    assert [r[-1] for r in bucketed.rows] == [BUCKET_0_90, BUCKET_91_180, BUCKET_NONE]
    assert counts == {BUCKET_0_90: 1, BUCKET_91_180: 1, BUCKET_NONE: 1}
    assert len(bucketed.rows) == 3 and len(bucketed.rows[0]) == len(detail.headers) + 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -m pytest tests/test_risky_inventory.py -k "cutoff or bucket" -q`
Expected: FAIL with ImportError (`compute_cutoff` not defined).

- [ ] **Step 3: Implement in `src/risky_inventory_engine.py`**

Add imports `from datetime import date, datetime, timedelta` (datetime already imported), and after `DETAIL_HEADERS`:

```python
MRP_LAST_SELL_COL = "MRP Last Sell Date"
BUCKET_COL = "Bucket"
BUCKET_0_90 = "0-90 Day"
BUCKET_91_180 = "91-180 Day"
BUCKET_NONE = "No Last Sell Date"
CUTOFF_DAYS = 90


def compute_cutoff(run_date: date) -> date:
    """The 0–90 / 91–180 dividing date: report run date + 90 days."""
    return run_date + timedelta(days=CUTOFF_DAYS)


def _as_date(value: Any):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def bucket_for(last_sell: Any, cutoff: date) -> str:
    """Bucket one MRP Last Sell Date value. Blank/non-date -> No Last Sell Date;
    on/before cutoff -> 0-90 Day; after cutoff -> 91-180 Day."""
    d = _as_date(last_sell)
    if d is None:
        return BUCKET_NONE
    return BUCKET_0_90 if d <= cutoff else BUCKET_91_180


def assign_buckets(detail: "DetailTable", cutoff: date) -> tuple["DetailTable", dict]:
    """Return a copy of ``detail`` with a Bucket column appended to every row,
    plus per-bucket counts. Row order is preserved."""
    li = detail.index[MRP_LAST_SELL_COL]
    counts = {BUCKET_0_90: 0, BUCKET_91_180: 0, BUCKET_NONE: 0}
    new_rows = []
    for row in detail.rows:
        b = bucket_for(row[li], cutoff)
        counts[b] += 1
        new_rows.append(list(row) + [b])
    bucketed = copy.copy(detail)
    bucketed.headers = detail.headers + [BUCKET_COL]
    bucketed.rows = new_rows
    return bucketed, counts
```

(`copy` and `Any` are already imported in this module.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -m pytest tests/test_risky_inventory.py -k "cutoff or bucket" -q`
Expected: PASS (3 tests).

- [ ] **Step 5: Remove the now-dead two-file code**

In `src/risky_inventory_engine.py` delete `KEY_COLUMNS`, `_row_key`, and `remove_duplicate_rows` (they are only used by the old two-file flow). Delete their tests from `tests/test_risky_inventory.py`: `test_already_cleaned_file_unchanged`, `test_cumulative_file_drops_exactly_the_90day_rows_in_order`, `test_same_material_different_batch_not_treated_as_duplicate`, `test_int_float_whitespace_and_date_normalisation_match`. (`_norm` stays only if still referenced; if now unused, delete it too.)

- [ ] **Step 6: Run the file's tests**

Run: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -m pytest tests/test_risky_inventory.py -q`
Expected: PASS (no ImportError from deleted symbols).

- [ ] **Step 7: Commit**

```bash
git add src/risky_inventory_engine.py tests/test_risky_inventory.py
git commit -m "feat(risky): add Last Sell Date bucketing, drop two-file dedup"
```

---

### Task 2: Build the committed PivotTable template asset

**Files:**
- Create: `scripts/build_risky_inventory_template.py` (one-time dev generator)
- Create: `src/templates/risky_inventory_template.xlsx` (committed binary asset, produced by the script)
- Test: `tests/test_risky_inventory.py`

**Interfaces:**
- Produces: a template workbook with sheets `Detail` (21-col header, no data) and `Summary` (a PivotTable over `Detail` with `Bucket` as a page filter, `refreshOnLoad=1`, empty/scrubbed cache).
- Consumes: the golden file `Risky Inventory June 24 P2 - 90D.xlsx` (dev machine only; NOT committed).

**Context for the implementer:** openpyxl cannot create a PivotTable, but it preserves an existing one through load→save (verified). The generator derives the template from the golden file, inserts a `Bucket` cache/pivot field at index 20 (before the grouped date fields, so no referenced field index shifts), scrubs all `sharedItems` + cached records to empty, and clears data rows. Excel rebuilds the cache from `Detail` on open via `refreshOnLoad`.

- [ ] **Step 1: Write the generator script**

Create `scripts/build_risky_inventory_template.py`:

```python
"""One-time generator for src/templates/risky_inventory_template.xlsx.

Derives the Risky Inventory PivotTable template from a golden export, adds a
'Bucket' page field, and scrubs all embedded supplier data (sharedItems +
cached records emptied; Excel rebuilds them on open via refreshOnLoad).

Usage:
    python scripts/build_risky_inventory_template.py ["path/to/golden 90D.xlsx"]
"""
from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.pivot.cache import CacheField, SharedItems
from openpyxl.pivot.table import FieldItem, PageField, PivotField
from openpyxl.utils import get_column_letter

GOLDEN = sys.argv[1] if len(sys.argv) > 1 else (
    r"C:/Users/melgh/Downloads/Risky Inventory June 24 P2 - 90D.xlsx"
)
OUT = Path(__file__).resolve().parents[1] / "src" / "templates" / "risky_inventory_template.xlsx"

EMPTY_RECORDS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>'
)


def main() -> None:
    wb = load_workbook(GOLDEN)
    ws1, ws2 = wb["Sheet1"], wb["Sheet2"]
    piv = ws2._pivots[0]
    cache = piv.cache

    ucol = ws1.max_column + 1            # 21 -> column U
    ws1.cell(1, ucol, "Bucket")

    # Insert Bucket as cache/pivot field 20 (after the 20 base columns, before the
    # grouped date fields). Append it as a page filter set to (All).
    cache.cacheFields.insert(20, CacheField(name="Bucket", sharedItems=SharedItems()))
    piv.pivotFields.insert(20, PivotField(axis="axisPage", showAll=False,
                                          items=[FieldItem(t="default")]))
    piv.pageFields.append(PageField(fld=20))

    # Rename sheets and repoint the cache at the (empty) Detail header.
    ws1.title, ws2.title = "Detail", "Summary"
    last = get_column_letter(ucol)
    cache.cacheSource.worksheetSource.sheet = "Detail"
    cache.cacheSource.worksheetSource.ref = f"A1:{last}1"
    cache.refreshOnLoad = True
    cache.recordCount = 0

    # Scrub embedded data: empty every field's shared items, then drop data rows.
    for cf in cache.cacheFields:
        cf.sharedItems = SharedItems()
    if ws1.max_row > 1:
        ws1.delete_rows(2, ws1.max_row - 1)

    buf = io.BytesIO()
    wb.save(buf)

    # Replace the cached records part with an empty one (no supplier data).
    OUT.parent.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename.endswith("pivotCacheRecords1.xml"):
                content = EMPTY_RECORDS
            zout.writestr(item, content)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Run the generator**

Run: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" scripts/build_risky_inventory_template.py`
Expected: prints `wrote .../src/templates/risky_inventory_template.xlsx`; the file exists.

- [ ] **Step 3: Write the template structure test**

Add to `tests/test_risky_inventory.py`:

```python
import zipfile
from pathlib import Path
from openpyxl import load_workbook as _load_wb
from src.risky_inventory_engine import TEMPLATE_PATH


def test_template_asset_is_valid_and_data_clean():
    assert Path(TEMPLATE_PATH).exists()
    wb = _load_wb(TEMPLATE_PATH)
    assert set(wb.sheetnames) == {"Detail", "Summary"}
    hdr = [wb["Detail"].cell(1, c).value for c in range(1, 22)]
    assert hdr[:20] == DETAIL_HEADERS and hdr[20] == "Bucket"
    piv = wb["Summary"]._pivots[0]
    names = [f.name for f in piv.cache.cacheFields]
    bucket_idx = names.index("Bucket")
    assert bucket_idx in [pf.fld for pf in piv.pageFields]       # Bucket is a page filter
    assert piv.cache.refreshOnLoad is True
    # No real supplier data committed in the cache definition.
    cdef = zipfile.ZipFile(TEMPLATE_PATH).read(
        "xl/pivotCache/pivotCacheDefinition1.xml").decode()
    assert "10001334" not in cdef
```

- [ ] **Step 4: Run the structure test**

Run: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -m pytest tests/test_risky_inventory.py::test_template_asset_is_valid_and_data_clean -q`
Expected: FAIL (TEMPLATE_PATH not yet defined in engine) → defined in Task 3 Step 1. If running Task 2 before Task 3, temporarily add `TEMPLATE_PATH` (next task Step 1) first; otherwise run after Task 3 Step 1.

- [ ] **Step 5: VERIFICATION GATE (manual, in Excel)**

Open `src/templates/risky_inventory_template.xlsx` in Excel. Confirm: it opens without repair prompts; the `Summary` sheet shows a PivotTable with `Bucket` in the Filters area alongside Description p. group / Brand Manager Desc / MRP Area. (It will be empty — the Detail sheet has no data yet.) **If Excel reports the file needs repair or the pivot is missing:** fall back to building the template by hand in Excel once (Detail header with Bucket as col 21; a PivotTable over Detail with the row/data fields from the spec and Bucket + the three existing page filters; PivotTable Options → Data → "Refresh data when opening the file"; save to the same path). The rest of the plan is unchanged.

- [ ] **Step 6: Commit**

```bash
git add scripts/build_risky_inventory_template.py src/templates/risky_inventory_template.xlsx
git commit -m "feat(risky): add scrubbed PivotTable template + generator"
```

---

### Task 3: Runtime template fill (`generate_excel`)

**Files:**
- Modify: `src/risky_inventory_engine.py`
- Test: `tests/test_risky_inventory.py`

**Interfaces:**
- Consumes: `assign_buckets` output (a `DetailTable` with the `Bucket` column), `TEMPLATE_PATH`.
- Produces: `generate_excel(bucketed: DetailTable) -> bytes` — a workbook whose `Detail` sheet holds the bucketed rows and whose `Summary` pivot is repointed at them with `refreshOnLoad=1`.

- [ ] **Step 1: Add `TEMPLATE_PATH` and remove the old summary code**

In `src/risky_inventory_engine.py` add near the top:

```python
from pathlib import Path

TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "risky_inventory_template.xlsx"
TEMPLATE_DETAIL_SHEET = "Detail"
TEMPLATE_SUMMARY_SHEET = "Summary"
```

Delete the old static-summary code: `SUMMARY_SHEET`, `SUMMARY_FILTERS`, `SUMMARY_GROUP_COL`, `SUMMARY_HEADER`, `GRAND_TOTAL_LABEL`, `_SUMMARY_VALUE_COLS`, `SUMMARY_COL_WIDTHS`, `SUMMARY_NUMBER_FORMATS`, `_num`, `build_summary`, `_write_summary_sheet`, `_to_excel_value`, `_write_detail_sheet`, and the **old** `generate_excel`. (The detail-writing happens inline in the new `generate_excel` below.)

- [ ] **Step 2: Write the failing test**

Add to `tests/test_risky_inventory.py`:

```python
import io as _io
from datetime import date
from src.risky_inventory_engine import generate_excel, assign_buckets, compute_cutoff


def test_generate_excel_fills_detail_and_keeps_pivot():
    rows = [
        _make_row(Material="A", **{"MRP Last Sell Date": datetime(2026, 8, 1)}),
        _make_row(Material="B", **{"MRP Last Sell Date": datetime(2026, 12, 1)}),
    ]
    detail = load_detail(_make_xlsx(rows))
    bucketed, _ = assign_buckets(detail, compute_cutoff(date(2026, 6, 24)))
    data = generate_excel(bucketed)

    # Pivot parts survive (real, refreshable PivotTable).
    parts = [n for n in zipfile.ZipFile(_io.BytesIO(data)).namelist() if "pivot" in n.lower()]
    assert len(parts) == 5

    wb = _load_wb(_io.BytesIO(data))
    det = wb["Detail"]
    assert det.cell(1, 21).value == "Bucket"
    assert det.max_row == 3                      # header + 2 rows
    assert det.cell(2, 21).value == "0-90 Day"
    assert det.cell(3, 21).value == "91-180 Day"
    piv = wb["Summary"]._pivots[0]
    assert piv.cache.refreshOnLoad is True
    assert piv.cache.cacheSource.worksheetSource.ref == "A1:U3"
```

- [ ] **Step 3: Run test to verify it fails**

Run: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -m pytest tests/test_risky_inventory.py::test_generate_excel_fills_detail_and_keeps_pivot -q`
Expected: FAIL (new `generate_excel` not yet defined / signature changed).

- [ ] **Step 4: Implement the new `generate_excel`**

In `src/risky_inventory_engine.py`:

```python
def generate_excel(bucketed: DetailTable) -> bytes:
    """Fill the committed PivotTable template with the bucketed detail and return
    the workbook bytes. The Summary pivot is repointed at the new data and set to
    refresh on open, so Excel rebuilds it from the Detail sheet."""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb[TEMPLATE_DETAIL_SHEET]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for r_off, row in enumerate(bucketed.rows):
        r = r_off + 2
        for c, value in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=value)
            fmt = bucketed.number_formats.get(c)
            if fmt:
                cell.number_format = fmt

    n = len(bucketed.rows)
    last_col = get_column_letter(len(bucketed.headers))   # 'U'
    piv = wb[TEMPLATE_SUMMARY_SHEET]._pivots[0]
    piv.cache.cacheSource.worksheetSource.ref = f"A1:{last_col}{n + 1}"
    piv.cache.recordCount = n
    piv.cache.refreshOnLoad = True

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
```

Ensure the module imports `from openpyxl import load_workbook` and `from openpyxl.utils import get_column_letter` (the file already imports `load_workbook`, `Workbook`, `get_column_letter`).

- [ ] **Step 5: Run the test + the template structure test**

Run: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -m pytest tests/test_risky_inventory.py -q`
Expected: PASS (bucketing, template structure, and fill tests all green).

- [ ] **Step 6: VERIFICATION GATE (manual, in Excel)**

In a Python shell, write a filled sample to disk and open it in Excel:

```python
from src.risky_inventory_engine import load_detail, assign_buckets, compute_cutoff, generate_excel
from datetime import date
d = load_detail(open(r"C:/Users/melgh/Downloads/Risky Inventory June 24 P2 - 180D.xlsx","rb"))
b,_ = assign_buckets(d, compute_cutoff(date(2026,6,24)))
open(r"C:/Users/melgh/Downloads/_ri_sample.xlsx","wb").write(generate_excel(b))
```

Open `_ri_sample.xlsx`: confirm the `Summary` PivotTable populates, the `Bucket` filter shows `0-90 Day` / `91-180 Day` / `No Last Sell Date`, and toggling it changes the totals. (If Excel prompts to enable/refresh, that's expected.)

- [ ] **Step 7: Commit**

```bash
git add src/risky_inventory_engine.py tests/test_risky_inventory.py
git commit -m "feat(risky): fill PivotTable template at runtime, drop static summary"
```

---

### Task 4: Rewrite the page (single upload + run-date)

**Files:**
- Modify: `pages/7_Risky_Inventory.py`
- Test: `tests/test_risky_inventory.py`

**Interfaces:**
- Consumes: `load_detail`, `compute_cutoff`, `assign_buckets`, `generate_excel`, bucket constants, `RiskyInventoryError`.

- [ ] **Step 1: Rewrite the page**

Replace the entire body of `pages/7_Risky_Inventory.py`:

```python
"""Risky Inventory — split one 0–180 day report into 0-90 / 91-180 buckets.

Upload the full 0–180 day report, pick the report run date, and the app buckets
each row by MRP Last Sell Date (cutoff = run date + 90 days), then builds a
workbook whose Summary sheet is a live Excel PivotTable filterable by Bucket.
"""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from src.risky_inventory_engine import (
    BUCKET_0_90,
    BUCKET_91_180,
    BUCKET_NONE,
    RiskyInventoryError,
    assign_buckets,
    compute_cutoff,
    generate_excel,
    load_detail,
)

st.title("Risky Inventory")
st.caption(
    "Upload the full 0–180 day report. Rows are split by MRP Last Sell Date into "
    "0-90 and 91-180 day buckets (cutoff = report run date + 90 days). The "
    "downloaded Summary sheet is a live Excel PivotTable you can filter by Bucket."
)

uploaded = st.file_uploader("Risky Inventory 0–180 day report (.xlsx)", type=["xlsx"], key="ri_file")
run_date = st.date_input("Report run date", value=date.today(), key="ri_run_date")
cutoff = compute_cutoff(run_date)
st.caption(f"Cutoff: 0-90 = MRP Last Sell Date on/before **{cutoff:%b %d, %Y}**; later → 91-180.")

if uploaded is None:
    st.session_state.pop("ri_processed_file", None)
    st.info("Upload the 0–180 day report to begin.")
    st.stop()

if st.button("Process file", type="primary"):
    st.session_state["ri_processed_file"] = uploaded.file_id
if st.session_state.get("ri_processed_file") != uploaded.file_id:
    st.info("Click **Process file** to split the report.")
    st.stop()


@st.cache_data(show_spinner=False)
def _process(file_bytes: bytes, cutoff_iso: str):
    detail = load_detail(io.BytesIO(file_bytes))
    bucketed, counts = assign_buckets(detail, date.fromisoformat(cutoff_iso))
    xlsx = generate_excel(bucketed)
    return bucketed, counts, xlsx


with st.spinner("Splitting the report and building the workbook..."):
    try:
        bucketed, counts, xlsx = _process(uploaded.getvalue(), cutoff.isoformat())
    except RiskyInventoryError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the file: {exc}")
        st.stop()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total rows", f"{len(bucketed.rows):,}")
c2.metric("0-90 Day", f"{counts[BUCKET_0_90]:,}")
c3.metric("91-180 Day", f"{counts[BUCKET_91_180]:,}")
c4.metric("No Last Sell Date", f"{counts[BUCKET_NONE]:,}")

st.subheader("Detail")
st.dataframe(
    pd.DataFrame(bucketed.rows, columns=bucketed.headers),
    use_container_width=True, hide_index=True,
)

st.subheader("Download")
st.caption("The Summary sheet is a live PivotTable — filter by Bucket (and the other filters) in Excel.")
st.download_button(
    "Download Risky Inventory Report",
    data=xlsx,
    file_name=f"Risky Inventory Report - {date.today():%B %d %Y}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
```

- [ ] **Step 2: Write the page smoke test**

Add to `tests/test_risky_inventory.py`:

```python
def test_page_renders_without_exception():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file("pages/7_Risky_Inventory.py", default_timeout=30).run()
    assert at.exception is None
```

- [ ] **Step 3: Run the test**

Run: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -m pytest tests/test_risky_inventory.py::test_page_renders_without_exception -q`
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add pages/7_Risky_Inventory.py tests/test_risky_inventory.py
git commit -m "feat(risky): single 0-180 upload with run-date split + pivot download"
```

---

### Task 5: Full integration verification

**Files:** none (verification only).

- [ ] **Step 1: Run the full test suite**

Run: `"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -m pytest tests/ -q`
Expected: all pass (existing Batch Quality / fill-rate / etc. tests + the rewritten Risky Inventory tests); only the pre-existing DB-gated skips remain.

- [ ] **Step 2: Confirm every page still loads**

Run:

```bash
"/c/Users/melgh/AppData/Local/Python/pythoncore-3.14-64/python.exe" -c "
from streamlit.testing.v1 import AppTest
import glob, os
bad=[]
for p in sorted(glob.glob('pages/*.py')):
    at=AppTest.from_file(p, default_timeout=30).run()
    if at.exception: bad.append((os.path.basename(p), str(at.exception)[:80]))
print('FAILURES:', bad or 'none')
"
```

Expected: `FAILURES: none`.

- [ ] **Step 3: Commit any final cleanup (if needed)**

```bash
git add -A
git commit -m "test(risky): full-suite + all-pages verification"
```

---

## Self-Review

- **Spec coverage:** single 0–180 upload + run-date picker (Task 4); bucketing by Last Sell Date incl. blank handling (Task 1); `Detail` + `Bucket` column and `Summary` PivotTable with Bucket page filter (Tasks 2–3); template+refresh mechanism, golden-derived + data-scrubbed (Task 2); engine API changes and dead-code removal (Tasks 1, 3); page UI (Task 4); tests + all-pages check (Tasks 1–5). All spec sections map to a task.
- **Placeholders:** none — every code step shows full functions/tests; the only manual steps are the two Excel verification gates, which are explicit checks, not code.
- **Type consistency:** `compute_cutoff(date)->date`, `bucket_for(value, date)->str`, `assign_buckets(DetailTable, date)->(DetailTable, dict)`, `generate_excel(DetailTable)->bytes`, `TEMPLATE_PATH` / `TEMPLATE_DETAIL_SHEET` / `TEMPLATE_SUMMARY_SHEET` are used consistently across Tasks 1–4. Bucket labels (`0-90 Day` / `91-180 Day` / `No Last Sell Date`) are referenced by constant everywhere.
- **Note:** `generate_excel` takes the bucketed `DetailTable` (run date already applied via the cutoff), a small simplification of the spec's `generate_excel(rows, run_date)` — no functional difference.
```
