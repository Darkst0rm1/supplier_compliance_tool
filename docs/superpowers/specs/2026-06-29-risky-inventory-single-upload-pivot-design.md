# Risky Inventory — Single 0–180 Upload + Real PivotTable

**Date:** 2026-06-29
**Status:** Approved design, pending implementation plan
**Component:** `pages/7_Risky_Inventory.py`, `src/risky_inventory_engine.py`, new template asset

## 1. Problem & Goal

The Risky Inventory page today takes **two** uploads — a 90-day report and a
"cumulative" 180-day report — and removes the 90-day rows from the 180-day set.
Two problems:

1. **The two-file premise doesn't match reality.** Inspecting the real exports
   (`Risky Inventory June 24 P2 - 90D/180D.xlsx`) shows the files are split by
   **`MRP Last Sell Date`**, not SLED offset, and the 90D/180D files are
   **disjoint** (0 overlapping rows) — the "180D" file holds only the 91–180 day
   items, so "subtract 90 from 180" removes nothing. The real boundary is the
   report run date + 90 days (run ~Jun 24 → cutoff Sep 22, the exact 90D/180D
   dividing line).
2. **The download has no real PivotTable.** The golden files carry a live,
   interactive Excel PivotTable on `Sheet2`. The current engine writes a
   *static* grid (`build_summary`) that only mimics the collapsed group-level
   view — dead cells, no expand / filter / refresh.

**Goal:** one upload of the full **0–180 day** report, split into buckets by
`MRP Last Sell Date`, output a workbook whose summary is a **real, interactive
PivotTable** the user can slice — matching the experience of their source files.

## 2. Input

- **One** `.xlsx` upload: the full 0–180 day Risky Inventory report. `Sheet1`
  holds the 20-column detail (`DETAIL_HEADERS`, unchanged). The single export is
  confirmed to contain both the 0–90 and 91–180 items together.
- **Report run date** — a date input on the page, defaulting to today. The
  user adjusts it to the actual run date if the file was downloaded earlier.
- **Cutoff** = run date + 90 days (the 0–90 / 91–180 dividing line).

## 3. Bucketing

Each detail row is assigned a `Bucket` from `MRP Last Sell Date`:

| Bucket | Rule |
|---|---|
| `0-90 Day` | Last Sell Date **≤ cutoff** (cutoff = run date + 90 days). Includes past dates. |
| `91-180 Day` | Last Sell Date **> cutoff**. |
| `No Last Sell Date` | Last Sell Date blank / missing / non-date. Pulled out separately. |

- Comparison normalizes datetime → date (`MRP Last Sell Date` values are
  datetimes; cutoff is a date). The boundary is **inclusive** on the 0–90 side
  (Last Sell Date == cutoff → `0-90 Day`), verified against the golden split
  (90D max Sep 22 = run+90; 180D min Sep 23).
- We **trust the export's 180-day ceiling** — no extra upper-bound filter is
  applied. Any row beyond 180 days (shouldn't occur) lands in `91-180 Day`.

## 4. Output Workbook

Two sheets, produced via a **template + refresh** mechanism (Section 5):

### `Detail`
All rows, original 20 columns **plus a `Bucket` column appended as the last
(21st) column**. Per-column number formats and header styling are preserved
from the uploaded file (as `load_detail` already captures them).

### `Summary`
One real, interactive **PivotTable** over the `Detail` sheet, mirroring the
golden pivot plus the new Bucket field:

- **Page filters:** `Bucket`, `Description p. group`, `Brand Manager Desc`,
  `MRP Area` (all default `(All)`).
- **Row fields (tabular, nested):** `Material Group Desc.` → `Material` →
  `Material Description` → `Batch` → `Batch Expiry Date`.
- **Data fields:** Sum of Quantity, Sum of Total Stock (`#,##0`), Sum of Value
  (`"$"#,##0`).
- Tabular layout (`compact=0`, `outline=0`), `refreshOnLoad=1` so Excel rebuilds
  it from `Detail` on open.

The `No Last Sell Date` rows are part of the same `Detail` sheet and selectable
via the Bucket page filter (satisfies the "separate / flagged" requirement
without a dead static sheet).

## 5. Template + Refresh Mechanism

openpyxl **cannot create** a PivotTable from scratch, but a load→save round-trip
**preserves** an existing one (verified: all 5 pivot parts survive, pivot
re-detected, and `refreshOnLoad` + source `ref` are editable). So:

### Template asset
A committed binary asset `src/templates/risky_inventory_template.xlsx`
containing:
- a `Detail` sheet with the 21-column header (20 source + `Bucket`), no data rows,
- a `Summary` sheet with the PivotTable from Section 4 wired to `Detail`,
  `refreshOnLoad=1`.

**How the template is built:** derived once from the user's golden file via a
one-time generator (a committed script / guarded helper) that:
1. renames `Sheet1` → `Detail`, `Sheet2` → `Summary`,
2. appends a `Bucket` column to the data sheet and a `Bucket` **page field** to
   the pivot (cacheField + pivotField + pageField, count bumps),
3. fixes the pivot `cacheSource` sheet name to `Detail` and clears data rows,
4. sets `refreshOnLoad=1`.

**Verification checkpoint (in the plan):** the generated template and a filled
sample output are **opened in Excel by the user** to confirm the pivot refreshes
and renders correctly. **Fallback:** if Excel rejects the programmatically built
template, the user builds the template once in Excel to this spec and commits it;
the runtime fill code is unchanged either way.

### Runtime fill — `generate_excel(detail_with_bucket, run_date) -> bytes`
1. Load the template with openpyxl.
2. Clear any `Detail` data rows (keep header); write all bucketed rows, applying
   the captured per-column number formats; `Bucket` in the last column.
3. Update the pivot cache: `worksheetSource.ref = A1:U{n+1}`,
   `recordCount = n`, `refreshOnLoad = True`.
4. Save to bytes. Excel rebuilds the pivot on open.

## 6. Page (`pages/7_Risky_Inventory.py`)

- One uploader: *"Risky Inventory 0–180 day report (.xlsx)"*.
- A `st.date_input("Report run date", value=today)`; show the derived cutoff
  (e.g. "0–90 = Last Sell Date on/before Sep 22, 2026").
- `Process` button (gated; persist processed-file id so reruns don't reset —
  same pattern used elsewhere in the app).
- Metrics: total rows, `0-90 Day`, `91-180 Day`, `No Last Sell Date` counts.
- A detail table showing the rows with the `Bucket` column (Streamlit-filterable).
- Download button → the template-filled workbook,
  `Risky Inventory Report - <Month DD YYYY>.xlsx`.
- A caption noting the downloaded `Summary` sheet is a live PivotTable
  (filterable by Bucket and the other page filters).

## 7. Engine Changes (`src/risky_inventory_engine.py`)

**Keep:** `RiskyInventoryError`, `DETAIL_SHEET`, `DETAIL_HEADERS`, `DetailTable`,
`load_detail` (including format/width/header-style capture).

**Remove (dead after this change):** `remove_duplicate_rows`, `KEY_COLUMNS`,
`_row_key`, `build_summary`, the old static-summary writers and `SUMMARY_*`
constants, and the old `generate_excel` signature.

**Add:**
- `MRP_LAST_SELL_COL = "MRP Last Sell Date"`, `BUCKET_COL = "Bucket"`,
  bucket label constants (`BUCKET_0_90`, `BUCKET_91_180`, `BUCKET_NONE`),
  `CUTOFF_DAYS = 90`.
- `compute_cutoff(run_date) -> date` (= run_date + 90 days).
- `assign_buckets(detail, cutoff) -> list[row]` returning rows extended with the
  `Bucket` value (and a count summary), preserving order.
- `TEMPLATE_PATH` resolution (package-relative).
- `generate_excel(detail_with_bucket, run_date) -> bytes` — the template fill of
  Section 5.

## 8. Edge Cases

- **Blank Last Sell Date** → `No Last Sell Date` bucket (Section 3).
- **Empty upload / no data rows** → valid workbook with an empty `Detail` and an
  empty-but-valid pivot (source range `A1:U1`); page shows zero counts.
- **Run date earlier/later than file** → user-controlled via the date picker;
  boundary rows follow the chosen cutoff.
- **Header mismatch** → `load_detail` already raises `RiskyInventoryError`.
- **Wrong sheet** → existing `Sheet1`-missing check still applies.

## 9. Testing

- `compute_cutoff` = run_date + 90.
- `assign_buckets`: inclusive 0–90 boundary, 91–180 above cutoff, blank →
  `No Last Sell Date`; counts correct; order preserved.
- `generate_excel`: output has `Detail` + `Summary`; `Detail` header ends with
  `Bucket`; rows written with formats; **pivot parts present** (5 parts),
  `refreshOnLoad=1`, cache `worksheetSource.ref` matches row count.
- Page smoke test (Streamlit `AppTest`) renders without exception; all 8 pages
  still load.
- The existing two-file tests are replaced by the above.

## 10. Risks

- **Template build fidelity** — the pivot-XML surgery is the main risk; mitigated
  by the Excel verification checkpoint and the user-built fallback (Section 5).
- **openpyxl pivot round-trip** — validated on the golden file; the grouped
  date cache fields (Months/Quarters/Years) are harmless and may be left as-is
  or stripped during template build.
- **Stale cached pivot values until open** — acceptable; `refreshOnLoad=1`
  rebuilds on open with no external-data prompt (worksheet-source cache only).
