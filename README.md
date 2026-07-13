# Supplier Documentation Compliance Tool — V1

Local tool that reconciles **SAP** purchase-order exports against **Portal**
upload exports and flags suppliers who skipped inbound documentation.

The core rule:

> If SAP has an Inbound Delivery for a PO but that PO is missing from the
> Portal export, the supplier did not submit the required documentation.

V1 is upload-only. No portal scraping, no MFA, no API.

---

## Install

```powershell
cd C:\Users\mohamed\supplier_compliance_tool
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python make_templates.py     # writes templates/sap_template.xlsx and portal_template.xlsx
```

## Run

```powershell
streamlit run app.py
```

Streamlit opens in your browser. Steps:

1. Upload the SAP export (`.xlsx`)
2. Upload the Portal export (`.xlsx`)
3. Pick the report **Year** and **Month**
4. Click **Generate Compliance Report**
5. Click **Download Excel report**

---

## Expected file shape

**SAP export columns** (exact names):

- PO Number, Vendor Number, Vendor Name, Warehouse, PO Status, Appointment Date, Delivery Date, Inbound Delivery, Inbound Delivery Status

`PO Status` codes: `C` = Closed, `A` = Approved, `P` = Processing.

**Portal export columns** (exact names):

- PO Number, Supplier Name, Upload Date

The portal export is filtered to the selected report month by `Upload Date`.
A portal cell may carry multiple POs separated by `,` `/` `;` newline or
spaces — they are split automatically.

PO numbers are normalized to text: trimmed, decimals like `.0` removed,
leading zeros preserved, whitespace stripped.

---

## Output workbook

One `.xlsx` with these sheets:

1. **Monthly Summary** — totals + compliance %
2. **Portal Export Data** — filtered + normalized portal rows
3. **SAP Export Data** — normalized SAP rows
4. **SAP Inbound Matched With Portal File** — compliant
5. **SAP Inbound Missing Portal File** — non-compliant (the main issue)
6. **Portal File But No SAP Inbound** — needs review
7. **Portal PO Not Found In SAP** — needs review
8. **Closed POs Review** — `PO Status = C`
9. **Processing POs Review** — `PO Status = P`
10. **Supplier Summary** — rollup by vendor
11. **Warehouse Summary** — rollup by warehouse

**Compliance % formula:**
`SAP Inbound POs With Portal File / Total SAP POs With Inbound Delivery`

---

## Layout

```
supplier_compliance_tool/
├── app.py                  Streamlit UI entry point
├── make_templates.py       Generates the two template files
├── requirements.txt
├── README.md
├── src/
│   ├── config.py           Column lists, status codes, labels
│   ├── normalizer.py       PO normalization + multi-PO split
│   ├── sap_importer.py     SAP loader/validator
│   ├── portal_importer.py  Portal loader/validator/month filter
│   ├── compliance_engine.py  Builds every report dataframe
│   └── report_generator.py   Writes/formats the .xlsx workbook
├── data/
│   ├── input/              (optional landing for input files)
│   └── output/             (optional download target)
└── templates/
    ├── sap_template.xlsx
    └── portal_template.xlsx
```

---

## Supplier Exceptions

Some suppliers are formally approved as **not required** to upload inbound
documentation. The **Supplier Summary** sheet flags these with an
**Exception Status** column: `Exception`, `Expected to upload`, or
`Not on tracker`.

A new **"Should Have Uploaded"** sheet + dashboard section is the chase-list:
suppliers with inbound deliveries this month, no uploads at all, and no
exception on file.

**Where the exception list comes from:** the Master Inbound Delivery
Compliance Tracker workbook, hand-maintained outside this tool. It's the
union of two sheets — `Tracker` rows with Compliance Status
`"NO -  Unable to Comply"`, and `POs received` rows hand-marked `EXEMPT` —
25 suppliers total after de-duplication. The tracker only names suppliers,
so matching against SAP vendor numbers is done by normalized name.

**Storage:** the list lives in a Neon Postgres table, `supplier_exceptions`.
`scripts/seed_supplier_exceptions.py` seeds/re-syncs it from the tracker
workbook and is idempotent (safe to re-run). **Once seeded, the database is
the source of truth** — editing the Excel workbook afterward has no effect
on the app. Day-to-day additions/removals happen in the "Manage Supplier
Exceptions" expander on the dashboard page.

**Informational only:** Exception Status does **not** change bill-back or
the compliance percentage — an exception supplier is still billed like any
other non-compliant supplier. This was an explicit business decision,
verified against real June data (bill-back unchanged at 59 suppliers /
$18,800).

**Configuration:** requires the `[postgres] dsn` secret, set in both
Streamlit Cloud (Settings → Secrets) and the local, git-ignored
`.streamlit/secrets.toml`. Without it the app still runs normally — it
shows an info banner and every supplier reads "Expected to upload" — it
never crashes.

---

## Future work (not in V1)

- Fine calculation
- Weekly report
- Email supplier follow-ups
- Dashboard charts
- User login
- Scheduled monthly reports
