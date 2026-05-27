# Build Log — Supplier Documentation Compliance Tool

**Date:** 2026-05-26
**Owner:** Mohamed (Tree of Life)
**Status at end of session:** Code complete + smoke-tested on real SAP data. Paused waiting for IT's portal export button.

---

## 1. Scope set at start of session

Pivoted from the earlier portal-scraping/Playwright/MFA/Entra approach (blocked) to an **upload-only** V1:

- User uploads a SAP export (`.xlsx`) and a Portal export (`.xlsx`)
- User picks a report month
- Tool compares POs by number, applies compliance rules, generates a multi-sheet Excel report
- **No** scraping, MFA, API, or fine calculation in V1
- Local tool only — Streamlit UI on `localhost:8501`

---

## 2. Initial build (V1 skeleton)

Created the project at `C:\Users\mohamed\supplier_compliance_tool\` with this layout:

```
supplier_compliance_tool/
├── app.py                       Streamlit UI
├── make_templates.py            Generates sample SAP + Portal Excel templates
├── requirements.txt             pandas, openpyxl, streamlit
├── README.md                    Install + run + sheet documentation
├── src/
│   ├── __init__.py
│   ├── config.py                Column names, status codes, semantic constants
│   ├── normalizer.py            PO normalization + multi-PO cell splitting
│   ├── sap_importer.py          SAP loader/validator
│   ├── portal_importer.py       Portal loader/validator + month filter
│   ├── compliance_engine.py     Builds every report dataframe
│   └── report_generator.py      Writes & formats the .xlsx workbook
├── data/{input,output}/         Optional file landing spots
└── templates/
    ├── sap_template.xlsx        Sample SAP file with real column shape
    └── portal_template.xlsx     Sample Portal file mirroring real portal page
```

Initial compliance rules implemented from the spec:

- SAP `Inbound Delivery` + Portal has PO → **Compliant**
- SAP `Inbound Delivery` + Portal missing PO → **Non-Compliant** (the key catch)
- Portal-only POs → Needs Review
- Closed (`C`) and Processing (initially `P`) get dedicated review sheets but aren't hidden
- Compliance % = `inbound POs with portal file / total SAP POs with inbound`

Installed missing dependency (`openpyxl`), ran `make_templates.py`, started Streamlit. App went live at `http://localhost:8501`.

---

## 3. Iterative discoveries from real data

### 3a. `export (6).xlsx` — first real SAP look

- 115 rows, 8 columns
- Columns: `PO Number, Confirmed PU Date, Est PU Date, Appt. Date, Delivery Date, Inbound Delivery Status, Inbound Delivery, Vendor Name`
- **Spec mismatch:** no `Vendor Number`, no `Warehouse`, no separate `PO Status` column
- The spec's "PO Status" codes (`C`/`A`/`P`) actually live in `Inbound Delivery Status`

### 3b. `export (8).xlsx` — canonical real shape

- 1 row, 10 columns (same as `export (6)` + `Plant` + `Vendor`)
- `Plant` is the Warehouse code (`2910`)
- `Vendor` is the Vendor Number (`70007031`)
- `Appt. Date` is the spec's `Appointment Date`

**Action taken:** added a column alias map in `sap_importer.py`:

```python
SAP_COLUMN_ALIASES = {
    "Plant": "Warehouse",
    "Vendor": "Vendor Number",
    "Appt. Date": "Appointment Date",
}
```

Also: if a SAP export has no `PO Status` column, the importer auto-copies `Inbound Delivery Status` into it (so the C/A/B routing still works).

Vendor Number and Warehouse made tolerant — absent columns get filled with blanks, with a soft warning surfaced in the Streamlit UI.

Two optional date columns (`Confirmed PU Date`, `Est PU Date`) added to the SAP Export Data sheet.

### 3c. `export (9).xlsx` — full-size SAP export (12,488 rows)

Profiled the real status code distribution:

| Code | Rows | Meaning |
|---|---|---|
| `C` | 8,917 | Closed |
| (blank) | 3,369 | No inbound delivery yet |
| `A` | 161 | Approved |
| **`B`** | **41** | **Processing / In-Progress** — spec called this `P` but real SAP uses `B` |
| `P` | 0 | Doesn't actually exist |

Also profiled date range: **May 2025 – Nov 2026** (18 months of history). This made the original "no SAP date filter" design wrong: with 18 months of POs mixing into a single month's compliance %, the number would be meaningless.

**Actions taken:**

- Changed `PO_STATUS_PROCESSING_CODES = {"B", "P"}` so both codes route to the Processing POs Review sheet (P kept as future-proofing)
- Added a "No Inbound Yet" sheet for the 3,369 blank-status rows (POs in scope but inbound not yet created)
- Added SAP-side month filtering (initially as a user-selectable date column — later changed; see §5)
- Verified column structure matches `export (8)` exactly across all 12,488 rows
- Confirmed PO Numbers are 100% unique (no line-item duplicates)
- Confirmed 5 warehouses present: `2910`, `2920`, `2925`, `2930`, `2935`

### 3d. Portal screenshots — `image001.png` and `image001 (1).png`

Portal page: `employee.treeoflife.com/Inbound/Delivery/list` ("ProConnect / Employee Portal"). Title: **Inbound Delivery File List**. Pagination shows 4,370 items total.

Real portal column shape (9 columns, not 3 as the spec said):

| Column | Example |
|---|---|
| `PO Number(s)` | `1000006767` or comma-separated multi-PO list |
| `File Name` | `PO 1000006767.xlsx` |
| `Uploaded By` | `fcko@kikkoman.com` (supplier-side) |
| `Supplier` | `KIKKOMAN SALES USA, INC.` |
| `File Status` | `Approved` / `Submitted` / `Invalid` |
| `Upload Date` | `5/22/2026, 5:08 PM` (datetime) |
| `Downloaded By` | `linda.vlasblom@treeoflife.com` (TOL staff) |
| `Download Date` | datetime, blank if not yet downloaded |
| `Invalid Comment` | rejection reason if `File Status = Invalid` |

**Actions taken:**

- Added portal alias map: `PO Number(s) → PO Number`, `Supplier → Supplier Name`
- Added the 6 new portal columns to the importer as optional (back-compat with the 3-column template)
- Confirmed multi-PO cells are real (comma-separated 10-digit POs) — existing `split_multi_po` logic already handles them
- Documented the spec for IT so the eventual export button matches the tool's expectations

---

## 4. File Status compliance logic

Designed on user confirmation that my recommendations should be used:

| File Status | Counts as | Where it lands |
|---|---|---|
| `Approved` | Compliant | SAP Inbound Matched With Portal File |
| `Submitted` | Compliant (supplier did their job on time) + flagged | SAP Inbound Matched **and** Pending TOL Review |
| `Invalid` | **Non**-Compliant | SAP Inbound Missing Portal File, with `Invalid Reason` column populated |
| (blank) | Compliant (back-compat for the minimal template) | Same as Approved |

**Sheets added because of this:**

- **Pending TOL Review** (new) — Submitted-status uploads, so TOL knows what's waiting
- **SAP Inbound Missing Portal File** now has two extra columns: `Portal File Status` and `Invalid Reason`

**Monthly Summary additions:**

- `Portal POs With Valid Upload (Approved/Submitted)`
- `Portal POs Marked Invalid`
- `Portal POs Pending TOL Review`
- `SAP Inbound POs Missing Portal File ...of which had an Invalid upload`

**Supplier and Warehouse Summary additions:** `Invalid Portal Uploads` column added to both rollups so vendors with rejected files surface in the per-supplier view.

---

## 5. SAP date filter — three-step decision

This took three back-and-forth turns to land:

1. **First design:** Streamlit dropdown lets the user pick `Delivery Date` / `Appointment Date` / `Confirmed PU Date` per run.
2. **User pushback:** "the date filter shouldn't be an option, it should be inside the code." → Removed the dropdown, hardcoded to `Delivery Date`.
3. **User clarification:** "I meant all 3 choices, pulled automatically without choosing one." → Replaced with a **UNION**: a PO is in scope if ANY of the date columns falls in the selected month.
4. **User catch:** "you forgot Est PU Date." → Added it. Final union is 4 columns.

**Final rule (hardcoded in `config.py`):**

```python
SAP_FILTER_DATE_COLUMNS = [
    "Delivery Date",
    "Appointment Date",
    "Confirmed PU Date",
    "Est PU Date",
]
```

A PO is in scope for the selected month if ANY of those four dates lies in that month.

**Impact on `export (9).xlsx` for May 2026:**

| Filter approach | POs in scope |
|---|---|
| Delivery Date only | 905 |
| 3-column union (no Est PU) | 1,002 |
| **Final 4-column union** | **1,147** |

Most of the additional POs are early-lifecycle (Est PU set but no inbound yet), which is correct — they belong in the "No Inbound Yet" sheet.

Monthly Summary prints the audit string `"Union of Delivery Date, Appointment Date, Confirmed PU Date, Est PU Date"` so anyone reading the report knows the rule.

---

## 6. Final output workbook — 13 sheets

1. **Monthly Summary** — totals + breakdowns + compliance %
2. **Portal Export Data** — filtered + normalized portal rows
3. **SAP Export Data** — month-scoped SAP rows
4. **SAP Inbound Matched With Portal File** — compliant POs
5. **SAP Inbound Missing Portal File** — non-compliant POs (with Invalid Reason if applicable)
6. **Pending TOL Review** — Submitted uploads waiting on TOL approval
7. **Portal File But No SAP Inbound** — needs review
8. **Portal PO Not Found In SAP** — needs review
9. **No Inbound Yet** — POs in scope this month with no SAP inbound delivery created yet
10. **Closed POs Review** — `Inbound Delivery Status = C`
11. **Processing POs Review** — `Inbound Delivery Status = B` (or `P`, future-proofing)
12. **Supplier Summary** — per-vendor rollup with `Invalid Portal Uploads`
13. **Warehouse Summary** — per-plant rollup with `Invalid Portal Uploads`

Every sheet has: dark-blue header row, frozen header (`A2`), auto-filter on every column, dates formatted `yyyy-mm-dd`, column widths auto-sized.

---

## 7. Validation done

- ✅ End-to-end smoke test against `export (9).xlsx` (12,488 SAP rows) + template portal — passed
- ✅ All 4 date columns folded into the union — verified 1,147 POs in scope for May 2026
- ✅ Status B routes correctly into Processing POs Review (10 rows)
- ✅ All 5 warehouses appear in Warehouse Summary
- ✅ File Status logic verified with the 3-row template (Approved + Submitted + Invalid)

---

## 8. Validation NOT done — outstanding for next session

1. **No real portal export has ever touched the tool.** Built from screenshots. Headers might drift slightly when IT ships the export button — expect to add aliases.
2. **User hasn't done a personal acceptance run on real SAP data** to confirm the closed/processing/no-inbound/supplier/warehouse sheets look correct to TOL's eye.
3. Awaiting IT — no concrete ETA was given.

---

## 9. When you come back

1. Drop the first real portal export in `Downloads/` and tell Claude — it will inspect columns and add any new aliases.
2. Confirm the `File Status` values match the assumed set `{Approved, Submitted, Invalid}`. If there are others, the engine's `PORTAL_VALID_STATUSES` and `PORTAL_PENDING_STATUSES` need updating.
3. Run a real-data report and accept (or flag) the output.

---

## 10. Deferred for later versions (out of V1 scope)

- Fine calculation
- Weekly reports
- Email follow-ups to non-compliant suppliers
- Dashboard charts in Streamlit
- User login
- Scheduled monthly runs
- Multi-month / historical comparison views

---

## How to run

```powershell
cd C:\Users\mohamed\supplier_compliance_tool
.\.venv\Scripts\Activate.ps1     # if you set up a venv; otherwise system Python works
streamlit run app.py
```

Then open `http://localhost:8501`.
