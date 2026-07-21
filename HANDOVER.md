# Handover — Supplier Compliance Tool

**Purpose of this document:** everything a new owner needs to run, maintain, and
eventually re-platform this tool without access to the original author.

**Repo:** `github.com/Darkst0rm1/supplier_compliance_tool`
**Live app:** Streamlit Community Cloud, entrypoint `app.py`, branch `main`
**Last updated:** 2026-07-16

Read [§1 Blockers](#1-blockers--read-this-first) before anything else. The code
is on GitHub and is not the risk. The credentials are the risk.

---

## 1. Blockers — read this first

The application **will not run** after handover until these three items are
transferred to company-owned accounts. All three currently sit on the original
author's personal accounts. Nothing else in this document matters until these
are resolved.

| # | Dependency | Used by | Where it lives now | If not transferred |
|---|---|---|---|---|
| 1 | **Neon Postgres** connection string | Column Variants (pages 2 & 3) | Author's Neon account. DSN in Streamlit Cloud Secrets UI + local `.streamlit/secrets.toml` (git-ignored, **not in the repo**) | Saved column layouts are lost; pages 2 & 3 lose the variant panel |
| 2 | **Anthropic API key** + billing | Batch Quality AI review (page 8) | `ANTHROPIC_API_KEY` env var. Billed to author's account | Page 8's AI review silently degrades to rules-only |
| 3 | **Streamlit Community Cloud** app | Whole app | Deployed from author's GitHub account | App goes offline; no one can redeploy |

**Actions required, in order:**

1. Create a company Neon (or any Postgres) database. Run the schema via
   `VariantStore.ensure_schema()` in `src/column_variants.py`. Put the **pooled**
   connection string (host contains `-pooler`, `sslmode=require`) into the new
   deployment's secrets as `[postgres] dsn`. See `.streamlit/secrets.toml.example`.
2. Issue a company Anthropic API key, set `ANTHROPIC_API_KEY` in the deployment
   environment. Note this is a **metered cost** — page 8 calls the API once per
   flagged issue group per run.
3. Transfer the GitHub repo to a company org, then redeploy from that account.

> **Note on secrets hygiene:** `.streamlit/secrets.toml` and `.env` are correctly
> git-ignored. Do not commit the real DSN or API key when re-pointing them.

---

## 2. What this tool is

Eight independent report generators behind one Streamlit sidebar. Despite four
pages being named "Dashboard", **this is not a dashboard product — it is an
Excel report generator with a web form in front of it.** Every page except
page 8's AI tab ends in a *Download .xlsx* button, and only pages 2, 3, and 4
draw charts at all.

That framing matters for the Power BI question — see [§7](#7-power-bi-migration-notes).

The pages share almost nothing. Each has its own engine module in `src/`, its
own input files, and its own business rules. **They can be migrated, rewritten,
or retired one at a time**, with the exception of pages 5 and 6, which read the
same two source files.

---

## 3. Page reference

Each page: what goes in, what rules apply, what comes out.

### Page 1 — Supplier Compliance Dashboard
- **Engine:** `src/compliance_engine.py`, `src/{sap,portal,receiving}_importer.py`, `src/config.py`
- **Inputs:** SAP export `.xlsx` + Portal export `.xlsx` + report Year/Month,
  **plus an optional Receiving Log `.xlsx`**
- **Output:** 13-sheet workbook (16 with a receiving log), plus one bill-back tab
  per offending supplier
- **Business question:** which suppliers failed to upload inbound delivery docs?

**Core formula:**
```
Compliance % = SAP inbound POs with a portal file / total SAP POs with an inbound delivery
```
POs with no inbound delivery yet are **excluded from the denominator** — they are
not yet eligible.

**Non-obvious rules (all in `src/config.py`):**
- **Month scoping is a UNION across four date columns** — `Delivery Date`,
  `Appointment Date`, `Confirmed PU Date`, `Est PU Date`. A PO is in scope if
  *any* one falls in the selected month. This is hardcoded deliberately and is
  **not** a user setting: compliance scope must be auditable, not tweakable. The
  union is printed into the Monthly Summary for that reason.
- **POs starting with `6` are dropped entirely**, from both SAP and Portal sides
  (`EXCLUDED_PO_PREFIXES`). Different PO type, not subject to portal docs.
- **Status codes are `C`/`A`/`B`, not `C`/`A`/`P`.** The written spec said `P`
  for processing; real exports use `B`. Both are accepted
  (`PO_STATUS_PROCESSING_CODES`). The status lives inside `Inbound Delivery
  Status` — there is no separate PO Status column in the real export.
- **File Status semantics:** `Approved` / `Received` / `Submitted` / blank all
  count as compliant. `Invalid` does **not**. `Submitted` also lands in a
  "Pending TOL Review" sheet — the supplier still gets credit.
- **Bill-back:** $200 per missing inbound document (`BILLBACK_FEE_PER_OCCURRENCE`).
  Uploaded-but-`Invalid` POs are **not** billed — the supplier attempted.

**Receiving Log — optional third input (`src/receiving_importer.py`)**

The dock's hand-kept log. It answers a question the portal structurally cannot:
not *was a file uploaded* but *did the paperwork match the goods on the truck*
(`Correct Batch` / `Correct BBD` / `Correct QTY`). Adds three sheets —
`Document Accuracy Exceptions`, `Portal vs Receiving Log`, `Receiving Log Data` —
and five columns to both the Supplier and Warehouse summaries.

- **Document accuracy never touches Compliance % or bill-back.** It is appended
  as its own block in the Monthly Summary. Folding a wrong BBD into the
  compliance figure would silently redefine a number people sign off on.
- **The accuracy denominator is POs the dock actually checked**, not all POs.
  Most of the log is unanswered; a supplier with no checks reads `n/a`, never
  `0%` or `100%`. Do not "fix" this by treating blank as a pass.
- **Coverage is genuinely partial and the UI says so.** In real June 2026 data
  only ~29% of rows carry an inbound-file answer. Treat the accuracy rate as a
  sample, not a supplier scorecard.
- **The workbook is messy by nature and the importer absorbs it:** the header row
  sits below a title banner (found by content, not position); header spacing
  drifts (`Y / N` vs `Y/N`, matched on a whitespace-stripped form); the Jan–Apr
  sheets predate the audit columns and are **skipped, not read as blank**;
  future-dated rows are empty appointment slots and fall out with the month
  filter; PO cells are hand-typed and may hold several POs split by `/ , & `.
- **Non-PO tokens are dropped and reported.** Real SAP POs are 7 or 10 digits,
  all numeric (verified against a 12,488-row export), so `RECEIVING_MIN_PO_DIGITS
  = 7` filters out carrier refs (`TR-34306`), supplier refs (`GHPO-23467`), and
  free text. The UI lists what it discarded rather than quietly under-reporting.
- **`Portal vs Receiving Log`** flags POs where the two sources disagree about
  whether a file exists. Each one is either a portal data problem or a dock
  data-entry problem — invisible from either source alone, and the closest thing
  to an acceptance test page 1 has until IT's portal export lands.

### Page 2 — Delivery Fill Rate
- **Engine:** `src/fill_rate_engine.py` (657 lines) — **no tests**
- **Input:** SAP/BW delivery fill-rate export
- **Output:** dashboard (3 chart/metric blocks) + Excel
- Column matching is **keyword-based**, not exact headers (`COLUMN_KEYWORDS`).
  Order matters: `total_short_amount` must precede `short_amount`.
- Uses the shared Column Variants panel → **requires Postgres**.

### Page 3 — Sales Order Fill Rate
- **Engine:** `src/sales_order_engine.py` (591 lines) — **no tests**
- **Input:** SAP/BW sales order fill-rate export
- **Output:** dashboard (4 chart/metric blocks) + Excel
- Same keyword-matching approach. Order matters: `unconfirmed_qty` before
  `unconfirmed_demand_amount`.
- Uses Column Variants → **requires Postgres**.

### Page 4 — Daily Short Report
- **Engine:** `src/daily_short_engine.py` (405 lines) — **no tests**
- **Input:** SAPUI5 Daily Short export (order-line level)
- **Output:** summary + three shorted analyses + Excel
- Fulfilment chain: `Order Qty → Confirmed → Total Delivery → Picked → Invoice`.
- Three analyses, all **Order-based**: unconfirmed (Order−Confirmed), shorted at
  outbound (Order−Delivered), shorted at invoicing (Order−Invoiced).
- The source workbook embeds a TEMPLATE below the data describing the report;
  the engine implements that template.

### Page 5 — Overstock Report
- **Engine:** `src/overstock_engine.py` (381 lines) — **no tests, but golden-file validated**
- **Inputs:** Materials inventory export + Last Sell / BDM master
- **Output:** 3 warehouse sheets — Mississauga / Calgary / Surrey
- **Status:** 257–259 of 260 golden rows reproduced exactly. Remaining diffs are
  hand-edits in the golden with no matching source row.

**Rules — all reverse-engineered from the business's finished workbook:**
1. Total stock (Unrestricted + Quality Inspection + Blocked) > 0
2. Plant in region — Mississauga `2910`; Calgary `2920`,`2925`; Surrey `2930`,`2935`
3. Storage Location = `1000` (main warehouse) **or** Customer Consignment (blank storage loc)
4. Material number does **not** start with `40` (packaging/display/promo)
5. Material matches a master `Product Number`
6. SLED present and **on/after** `report_date + 6 days`
7. Last-sell-by **on/before** `report_date + 7 days`
8. **Not** the RANA retail line (Brand Manager `SANDRA GAGANIARAS GB` **and**
   description starts `RANA`) — note "RANA FS" foodservice is **kept**
9. **Not** Sweet Street (description starts `SSD`)

### Page 6 — Donate / Dispose List
- **Engine:** `src/donate_dispose_engine.py` (330 lines) — **no tests, golden-file validated**
- **Inputs:** same two files as page 5
- **Output:** 3 region sheets. All three reproduced exactly (48 / 92 / 19 rows)
- **The mirror image of Overstock.** Same source, flipped date window: keeps
  stock at/near/past expiry. Key differences from page 5: **no** storage-location
  restriction (all locations), **no** last-sell-date filter — the SLED cutoff
  (`report_date + SLED_CUTOFF_OFFSET_DAYS`) alone defines the window. Sorted by
  SLED ascending.
- Intentionally **does not import** from `overstock_engine` — kept self-contained
  so the two sets of rules can drift independently. Do not "DRY" these together.

### Page 7 — Risky Inventory
- **Engine:** `src/risky_inventory_engine.py` (410 lines) — **12 tests, golden-file validated**
- **Inputs:** 90-day report + cumulative 180-day report
- **Output:** 4 sheets (90D/180D Detail + Summary), byte-exact to golden
- **Scope is deliberately tiny.** Automates exactly one manual step: removing
  from the 180-day detail the rows already present in the 90-day detail. Nothing
  is renamed, recalculated, filtered, or added. Matching compares **16 key
  fields**, not Material alone.
- Sheet2 is a real Excel PivotTable, not a static table.

### Page 8 — Batch Quality Analysis
- **Engine:** `src/batch_quality/` (9 modules)
- **Input:** SAP receiving-history export
- **Output:** 4 tabs + Excel. **34 tests** — best-covered module
- **Uses the Anthropic API.** Rule-based detection runs first, then an AI agent
  reviews flagged issue groups automatically. Falls back to rules-only if the
  key is absent — *silently*, which is a trap: a missing key looks like "no AI
  findings", not an error.
- Env vars: `ANTHROPIC_API_KEY`, `ANTHROPIC_MODEL`, `BATCH_QUALITY_MAX_AI_ISSUES`,
  `BATCH_QUALITY_AI_CONCURRENCY`. An OpenAI-compatible endpoint can be swapped in
  via `BATCH_QUALITY_LLM_BASE_URL` / `BATCH_QUALITY_LLM_API_KEY`.

---

## 4. Cross-cutting design principles

These were learned the hard way. Violating them has broken this tool before.

- **Compliance scope is not configurable.** If a setting changes the compliance
  % or who gets flagged, it is a hardcoded business rule surfaced *in the report*
  — so the rule is auditable, not the setting. (The 4-column date union is the
  case study: it started as a dropdown and was deliberately removed.)
- **Real exports never match the spec.** SAP spec said PO Status was its own
  column; it isn't. Portal spec said 3 columns; real exports have 9. Status spec
  said `P`; data says `B`. **Always inspect a real file before trusting headers.**
  Importers use alias maps so header drift degrades gracefully instead of erroring.
- **Be liberal about optional columns.** Only `PO Number` and `Inbound Delivery`
  are hard-required on the SAP side. Everything else fills blank so partial
  exports still load.
- **PO normalization:** trimmed, `.0` decimals stripped, leading zeros preserved,
  treated as text throughout. Portal cells may hold multiple POs separated by
  `,` `/` `;` newline or space — split automatically.
- **"All of them" means UNION**, not separate views per filter.

---

## 5. Tests and validation

118 tests total, via `pytest`. **Coverage is very uneven — know where you are safe:**

| Module | Tests | Also golden-validated? |
|---|---|---|
| `batch_quality` | 34 | no |
| `receiving_log` | 28 | no |
| `column_variants` | 24+1 | no |
| `billback` | 14 | no |
| `risky_inventory` | 12 | **yes** |
| `po_exclusion` | 3 | no |
| `compliance_engine` | 2 | no |
| `fill_rate_engine` | **0** | no |
| `sales_order_engine` | **0** | no |
| `daily_short_engine` | **0** | no |
| `overstock_engine` | **0** | **yes** |
| `donate_dispose_engine` | **0** | **yes** |

**Read this table as risk.** `compliance_engine.py` is 593 lines of the most
business-critical logic in the tool and has 2 tests. The three engines with zero
tests *and* no golden file (pages 2, 3, 4) are the ones where a regression would
go unnoticed.

The "golden file" pattern is this tool's real safety net: the business supplied
a finished workbook, and the engine is validated by reproducing it row-for-row.
**Keep the golden files.** They are what makes pages 5, 6, and 7 safe to change
— or to re-platform.

---

## 6. Running it

```powershell
cd supplier_compliance_tool
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python make_templates.py        # writes templates/*.xlsx
streamlit run app.py            # http://localhost:8501
```

Tests: `pip install -r requirements-dev.txt && pytest`

Runtime deps: `pandas`, `openpyxl`, `streamlit`, `plotly`, `psycopg[binary]`,
`anthropic`, `pydantic`.

Without a Postgres DSN the app still runs — pages 2 and 3 just lose the Column
Variants panel. Without `ANTHROPIC_API_KEY`, page 8 runs rules-only.

**Also read `BUILD_LOG.md`** — chronological narrative of why decisions were made.
Design docs for two features are in `docs/superpowers/`. A user-facing training
guide is at `docs/dashboard_training_guide.pdf`.

> ⚠️ `README.md` is **out of date** — it documents V1 only (page 1, 11 sheets),
> predates pages 2–8, still says status code `P`, and omits bill-back and the
> PO-prefix exclusion. Trust this document over the README.

---

## 7. Power BI migration notes

If the goal is re-platforming to Power BI, the honest split:

**Maps well (genuinely better in Power BI):** pages 2, 3, 4. These are
aggregate-and-chart workloads — KPIs, rates, group summaries. They also have
zero tests today, so nothing is lost from the safety net. Slicers and
drill-through come free.

**Does not map:** pages 1, 5, 6, 7. These produce formatted multi-sheet Excel
deliverables that people sign off on. Power BI has no real Excel-writing story —
"Export to Excel" gives a flat dump, not a 13-sheet workbook with per-supplier
bill-back tabs. Paginated reports can approximate it but need Premium-Per-User or
Fabric capacity. **Pages 5 and 6 are validated byte-for-byte against golden files;
that guarantee cannot survive the move.**

**Cannot map at all:** page 8. Power BI does not host an LLM agent.

**The structural blocker:** today a user picks files off their desktop and hits
go. Power BI has no equivalent — a report viewer cannot hand a file to a dataset.
Files must land somewhere Power BI can reach on a schedule (SharePoint, OneDrive,
a database). That is a *process* change for whoever runs these reports, and it is
usually what kills these migrations.

**Recommended approach — split by output type, not by page:**
- Dashboards (2, 3, 4) → Power BI
- Excel deliverables (1, 5, 6, 7) → **Excel + Power Query**. Power Query's M
  language runs in Excel too, so the company learns *one* new language and keeps
  native Excel output with no new licensing.
- Page 8 → stays Python, or is retired.

**Migration technique:** do not rewrite blind. **Keep the Python as an oracle** —
run old and new side by side on the same input and diff until they match. The
golden files for pages 5, 6, 7 are already exactly this harness, which makes
those three the *safest* to port despite looking hardest. Migrate one page at a
time; turn nothing off until its replacement diffs clean.

---

## 8. Open items inherited

1. **Page 1 has never had a real acceptance run.** The portal export button was
   still being built by IT as of the last update. Column headers may drift from
   the 9-column shape the importer expects — expect to add aliases to
   `PORTAL_COLUMN_ALIASES` when the first real export lands. Confirm File Status
   values match `{Approved, Received, Submitted, Invalid}`; extend
   `PORTAL_VALID_STATUSES` / `PORTAL_PENDING_STATUSES` if not.
   **Partial substitute available now:** the receiving log's `Inbound File Y/N`
   is an independent human record of the same fact. Run a month with both files
   and read the `Portal vs Receiving Log` sheet — it is the only cross-check page
   1 has until the portal export exists. Coverage is ~29% of dock rows, so it
   narrows the risk rather than closing it.
2. **SAP Master Data exports may contain two stacked extracts** — the first batch
   carries IB Delivery as a number, the second as text with a leading zero, and
   the lower row sometimes has *more* data. Merge before doing lookups.
3. **The app is public.** Anyone with the Streamlit URL can use it. A viewer
   allowlist exists in Streamlit settings if access needs restricting.
4. **`README.md` needs rewriting** to match the current 8-page tool.
5. Deferred features, never built: fine calculation, weekly reports, email
   supplier follow-ups, user login, scheduled monthly reports.
</content>
