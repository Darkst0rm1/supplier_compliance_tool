# Non-Compliant Bill-Back — Design

**Date:** 2026-06-12
**Status:** Approved (design), pending implementation
**Page:** Supplier Compliance Dashboard (`pages/1_Supplier_Compliance_Dashboard.py`)

## Goal

Charge suppliers a bill-back fee for inbound POs where **no portal documentation file was ever uploaded**. Output one Excel tab per non-compliant supplier listing the offending PO numbers and the total amount to be charged, delivered inside the existing compliance workbook (no separate button or upload).

## Charge rule

- **Fee:** $200 (USD) per occurrence.
- **Occurrence:** one unique SAP PO that has an inbound delivery but **no portal file was ever submitted**.
- **Reason label:** "Missing Inbound Document".
- **Excluded from billing:** POs where the supplier uploaded a file that TOL later marked **Invalid/rejected**. The supplier attempted, so these are not charged. (They remain visible in the existing `SAP Inbound Missing Portal File` sheet with their rejection reason.)

## Billable set derivation

Source: the existing `missing` bucket in `compliance_engine.build_report`
(`missing = sap_unique[Has Inbound & ~Portal Match]`).

Billable = `missing` rows where `Portal Invalid Match == False`
(equivalently, `_missing_issue_text(row) == "No portal file was submitted."`).

Each billable row is already one unique PO (`sap_unique` is de-duplicated on
`Normalized PO Number`), so row count = occurrence count.

## Config additions (`src/config.py`)

```python
BILLBACK_FEE_PER_OCCURRENCE = 200          # USD per missing inbound document
BILLBACK_REASON = "Missing Inbound Document"
```

## Output: one tab per supplier

Grouping key: `Vendor Number` (stable) with `Vendor Name` for display.
Suppliers ordered by total charge **descending** (biggest offenders first).

Each supplier tab columns:

| PO Number | Warehouse | PO Status | Appointment Date | Delivery Date | Inbound Delivery | Charge Reason | Charge (USD) |
|---|---|---|---|---|---|---|---|

Final row of each tab is a total row:
`PO Number = "TOTAL — <N> occurrences"`, `Charge (USD) = N * 200`, other cells blank.

### Sheet naming

- Prefix `BB-` so all bill-back tabs group together and are identifiable.
- Append a sanitized, truncated vendor name.
- Excel constraints: max 31 chars; strip the illegal characters `: \ / ? * [ ]`.
- On collision (after truncation), append a numeric suffix (`-2`, `-3`, …) and
  re-truncate to stay within 31 chars.

No summary/index tab — strictly one tab per supplier, per the user's explicit choice.
The grand total is not aggregated anywhere; each supplier's total lives in its own tab.

## Integration points

1. `src/config.py` — add the two constants above.
2. `src/compliance_engine.py` — add a helper that takes the `missing` dataframe,
   filters to never-uploaded rows, and returns an **ordered** `dict[str, DataFrame]`
   of `{sheet_name: supplier_tab_df}`. Merge these entries into the dict returned by
   `build_report` (appended after the existing sheets so they render last).
3. `src/report_generator.py` — no change needed; `generate_workbook` already iterates
   the sheet dict and formats every sheet. The total row will render as a normal row.
4. `pages/1_Supplier_Compliance_Dashboard.py` — no new button. Optionally surface the
   number of suppliers billed and total dollar amount as a caption/metric after the
   report is generated (nice-to-have, not required for the export to work).

## Edge cases

- **No billable POs:** emit no bill-back tabs (or rely on the existing empty-sheet
  placeholder behavior). The workbook is otherwise unchanged.
- **Blank/missing Vendor Name:** fall back to `Vendor Number`; if both blank, use
  `"Unknown Supplier"` and still bill (the POs are real).
- **Duplicate sanitized names** across different vendors: handled by the numeric-suffix
  dedupe described above.

## Out of scope

- The `Discrepancy` portal status (deferred earlier — not added).
- Any persistence, emailing, or invoicing of the bill-back. This produces the Excel
  tabs only.
