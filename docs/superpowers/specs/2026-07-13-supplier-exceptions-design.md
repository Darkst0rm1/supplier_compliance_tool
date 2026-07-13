# Supplier Exceptions — Design

**Date:** 2026-07-13
**Status:** Approved (pending spec review)

## Problem

The Supplier Compliance Dashboard treats every supplier identically: if a SAP PO
has an inbound delivery and no portal upload, the supplier is non-compliant and
gets billed. In reality a group of suppliers has been granted an **approved
exception** — they are not required to upload inbound documentation at all.
Ops tracks these by hand in the *Master Inbound Delivery Compliance Tracker*
workbook, and nothing in the app knows about them.

Two things are missing:

1. The **Supplier Summary** sheet cannot show whether a supplier is an exception.
2. There is no way to answer the operational question: *which suppliers uploaded
   nothing at all this month even though they were expected to?*

## Scope

**In scope**

- An `Exception Status` column on the **Supplier Summary** sheet.
- A new **Should Have Uploaded** sheet + dashboard section listing suppliers who
  uploaded nothing despite being expected to.
- A Neon-backed `supplier_exceptions` table, seeded from the tracker, editable
  in-app.

**Explicitly out of scope — do not change**

- **Bill-back is untouched.** Exception suppliers are still billed exactly as
  they are today. The status is informational only.
- **Compliance percentage is untouched.** Exception suppliers stay in the
  denominator.

This was a deliberate decision: acting on the exception status moves real
invoices, and the user chose to see the data before changing who gets charged.
A future change can flip this, and the design keeps that door open by putting the
status on every SAP row rather than only in the summary rollup.

## What counts as an exception

The exceptions list is the **union** of two lists in the tracker
(`Master Inbound Delivery  Compliance Tracker_April 2025.xlsx`):

1. `Tracker` sheet, `Compliance Status == "NO -  Unable to Comply"` — 24
   suppliers. The `Summary` sheet relabels this row *"Approved exceptions"*.
2. `POs received` sheet, unnamed 4th column `== "EXEMPT"` — 3 suppliers.

The two lists overlap on 2 suppliers (Bothwell Cheese, Dare/Lesley Stowe), so the
union is **25 suppliers**, of which only Lundberg Family Farms is unique to list 2.

Note the exact spelling `"NO -  Unable to Comply"` — it contains a **double
space** after the dash, and differs from the `Summary` sheet's wording
(`"NO - Unable to comply - Approved exceptions"`). Match on the `Tracker` sheet's
literal value; do not retype it.

## The join problem

**The tracker's supplier identifiers do not join to SAP.** Its `Supplier #`
column holds values like `G61`, `B50`, `491`; SAP's `Vendor Number` is an 8-digit
code like `70007212`. Overlap across all 73 vendors in the June export: **zero**.

Therefore the join is **by supplier name**, normalized:

```
normalize(s) = " ".join(s.upper().replace(".,'-()&" -> " ").split())
```

Even normalized, only **33 of 72** SAP vendor names match a tracker name. The
rest are either genuinely absent from the tracker or are not suppliers at all —
`AMERICOLD TACOMA`, `CONGEBEC MISSISSAUGA`, `CJ LOGISTICS` and
`CONESTOGA COLD STORAGE` are 3PL warehouses.

This forces a **three-state** status rather than a boolean:

| Status | Meaning |
|---|---|
| `Exception` | On the exceptions list — not required to upload. |
| `Expected to upload` | On the tracker, not an exception. |
| `Not on tracker` | No tracker entry. Treated as expected to upload, but visibly flagged. |

`Not on tracker` exists so the tracker's incompleteness stays visible. Folding
these into `Expected to upload` would make a supplier who is missing from the
tracker by oversight indistinguishable from one who is genuinely non-compliant,
and would leave the 3PL warehouses permanently sitting in the chase-list looking
like delinquent suppliers.

## Architecture

### `src/supplier_exceptions.py` (new)

Neon-backed store. **Mirrors `src/column_variants.py` exactly** — same psycopg3
pattern, same `st.secrets["postgres"]["dsn"]` → `DATABASE_URL` env fallback, same
graceful degradation.

```
supplier_exceptions
  id              serial primary key
  supplier_name   text not null          -- as written in the tracker, for display
  normalized_name text not null unique   -- the join key
  vendor_number   text                   -- nullable; hardens matching over time
  reason          text not null          -- 'Unable to Comply' | 'EXEMPT mark' | 'Manual'
  added_by        text
  added_at        timestamptz default now()
```

Public API:

- `load_exceptions() -> dict[str, ExceptionRecord]` keyed by `normalized_name`
- `add_exception(name, reason, vendor_number=None)`
- `remove_exception(normalized_name)`

`vendor_number` is nullable because the tracker cannot supply it. It is populated
opportunistically: when a name match against SAP is confident, the vendor number
is recorded so subsequent runs match exactly instead of by name.

**Fail-open, not fail-closed.** If the DB is unreachable, `load_exceptions()`
returns `{}` and the report still generates — every supplier shows
`Unknown (DB unavailable)` in the Exception column. Because bill-back and
compliance % are unaffected by this feature, an outage cannot wrongly excuse a
supplier from a charge; it only removes an annotation. This mirrors the
Standard-only fallback the column-variants feature already ships.

### `src/compliance_engine.py` (modified)

`build_report(sap_df, portal_df, report_year, report_month, exceptions=None)`.

`exceptions` is a `dict[normalized_name -> ExceptionRecord]`; defaulting it to
`None` keeps every existing caller and test working unchanged.

- Annotate `sap_valid` with `Exception Status` (the three states above), derived
  from `Vendor Name` via `normalize()`, falling back to `Vendor Number` when the
  record carries one.
- `_supplier_summary()` gains an `Exception Status` column.
- New `_should_have_uploaded(sap_unique)` builds the new sheet.

Bill-back (`_billback_sheets`) and the compliance percentage are **not touched**.

### The "Should Have Uploaded" sheet

One row per supplier where **all** hold:

- at least one SAP PO with an inbound delivery in the report month, **and**
- **zero** portal uploads of any kind — no valid upload *and* no Invalid upload
  (a rejected upload still means the supplier knows the process exists), **and**
- `Exception Status != "Exception"`.

Columns: `Vendor Number`, `Vendor Name`, `Exception Status`,
`Inbound POs Expected`, `Portal Uploads` (always 0), `Bill-Back Total`.
Sorted by `Inbound POs Expected` descending — biggest offenders first.

The "uploaded *nothing*" test is deliberately stricter than "missing *some*". A
supplier who uploaded 9 of 10 has a working process with a gap; one who uploaded
0 of 10 does not know the process exists. The existing bill-back tabs already
cover partial cases, so this sheet earns its place only by isolating the
total-failure population.

### `pages/1_Supplier_Compliance_Dashboard.py` (modified)

- Load exceptions and pass them to `build_report`.
- New dashboard section: **Should Have Uploaded** — headline count + table.
- A collapsed `st.expander("Manage Supplier Exceptions")` with the current list,
  an add form, and a remove control. Minimal by design; the tracker workbook
  remains where Ops does bulk editing.
- If the DB is unavailable, show an `st.info` explaining the column reads
  `Unknown` — never a crash.

### `scripts/seed_supplier_exceptions.py` (new)

One-off, **idempotent** (`ON CONFLICT (normalized_name) DO NOTHING`) import of
the 25 suppliers from the tracker workbook. Takes the workbook path as an
argument. Re-runnable when the tracker changes.

## Testing

`tests/test_supplier_exceptions.py`:

- `normalize()` — punctuation, casing, whitespace collapsing.
- Union logic — 24 ∪ 3 = 25, with the 2 overlaps deduplicated. Guards against a
  regression where the double-space in `"NO -  Unable to Comply"` is "corrected".
- Three-state classification, including a vendor present in neither list.
- `build_report` **without** `exceptions` produces byte-identical sheets to today
  — the safety net for "we changed nothing else".
- Bill-back totals and compliance % are **unchanged** when exceptions are passed.
  This is the test that enforces the out-of-scope promise.
- "Should Have Uploaded" excludes exception suppliers, excludes suppliers with an
  Invalid upload, and includes a zero-upload non-exception supplier.
- Empty-frame safety per `feedback_pandas3_empty_frames` — a month with no SAP
  rows must not crash.
- DB integration tests gated on `TEST_DATABASE_URL`, as the column-variants tests
  already are.

## Risks

- **Name matching is the weak point.** 39 of 72 SAP vendors will land in
  `Not on tracker` on the June data. That is expected and visible, not a bug —
  but if Ops reads it as noise, the column loses value. The `vendor_number`
  backfill is the long-term fix.
- **Tracker/DB drift.** Once seeded, the DB is the source of truth. Editing the
  Excel does not update the app. The re-runnable seed script is the escape hatch;
  the in-app editor is the intended path.
