# Shared Column Variants — Design Spec

- **Date:** 2026-06-19
- **Status:** Draft for review
- **Repo:** supplier_compliance_tool
- **Supersedes:** the reverted "variance profiles" feature (commits V6–V10)

## 1. Summary

Add an SAP-style **Variant** picker above the main detail table on the Delivery
Fill Rate and Sales Order Fill Rate dashboards. A *variant* is a named, ordered
**subset of table columns** — it controls **which columns the table shows and in
what order**, nothing else (no filter values).

- **"Standard"** is built in and read-only = every column, in natural order.
- User-created variants are **shared**: stored in a hosted **Postgres** table and
  **editable by every user** (the app has no login, so there is no concept of a
  private variant).
- The active variant drives **both the on-screen grid and the matching sheet in
  the Excel export**.

This replaces the reverted variance-profiles build, which was over-engineered
(multi-tenant SQLite, permissions, audit log), saved the wrong thing (filter
values against fixed widgets), and was buggy.

## 2. Goals

- A `Standard ▾` dropdown above each target table listing `Standard` + all shared
  variants for that report.
- Full CRUD on shared variants, available to **all** users: **create, view,
  update, rename, delete**.
- Variants **persist across Streamlit Cloud restarts and redeployments** (hosted
  Postgres, not ephemeral storage).
- Variants are **separated by report key** (one table, a `report_key` column).
- The active variant projects columns onto the displayed table **and** the
  corresponding Excel sheet.
- The dashboards keep working (Standard only) if the database is unreachable.

## 3. Non-goals (explicitly out of scope)

- Authentication / per-user or private variants / ownership / permissions / audit log.
- Saving **filter values** — variants are columns-only.
- A "set as default" variant (Standard is always the landing view). Easy to add later.
- `localStorage`, local JSON, or local SQLite persistence.
- Page 1 (Supplier Compliance — generate-and-download, no on-screen grid) and
  Page 4 (Daily Short Report).
- Re-columning the secondary sheets/tabs (summaries, Top 10, Clean Data, Raw
  Preview). The variant applies to the **one detail table per page** listed below.

## 4. Scope — where it appears

| Dashboard | Tab / table | `report_key` |
|---|---|---|
| Page 2 — Delivery Fill Rate | **Shortage Report** (`shortage_df`) | `delivery_shortage` |
| Page 3 — Sales Order Fill Rate | **Unconfirmed Demand Report** (`unconfirmed_df` / `view`) | `sales_order_unconfirmed` |

Both tables are wide "all columns of the cleaned data, filtered to problem rows"
grids; the available columns depend on the uploaded file, so variants reference
columns **by name** and are intersected with the columns actually present at
render time.

## 5. Behavior / UX

Panel rendered inside the tab, directly above the grid:

```
Variant: [ Standard ▾ ]            ● unsaved
Columns shown: [product ✕] [plant ✕] [short_amount ✕] [+ add column…]
[ 💾 Save ] [ 📑 Save as… ] [ ✏️ Rename ] [ 🗑️ Delete ]   ↩ Reset to Standard
──────────────────────────────────────────────────────────────────
 product | plant | short_amount | …rows…
```

- **Dropdown** options: `Standard` first, then shared variant names (sorted,
  case-insensitive).
- **Columns shown** is a `st.multiselect`; **selection order = display order**, so
  show/hide and reorder happen in one control. Standard seeds it with all columns.
- **Dirty marker** (`●`/`* unsaved`) shows when the current selection differs from
  the loaded variant's stored columns (or, on Standard, differs from "all columns").
- **Buttons:**
  - **Save** — enabled only when a **shared variant** is active **and** dirty;
    `UPDATE` its `columns`.
  - **Save as…** — always available; name input → `INSERT` a new shared variant.
    Used to fork Standard or any variant.
  - **Rename** — shared active only; `UPDATE name`.
  - **Delete** — shared active only; `DELETE`, then fall back to Standard.
  - **Reset to Standard** — reselect Standard / restore all columns.
- **Standard is read-only:** Save/Rename/Delete are disabled while it is active.
- Success/error shown as a flash message; the variant list refreshes after any
  mutation and the page reruns.

### Validation rules

- Name is trimmed, non-empty, ≤ 120 chars.
- The name **"Standard"** is reserved (case-insensitive): cannot be created,
  renamed-to, or deleted.
- A variant must have **≥ 1 column**; empty selections cannot be saved.
- Duplicate names are rejected **case-insensitively** (`Mgmt View` == `mgmt view`),
  enforced by the `LOWER(name)` unique index and a matching app-level pre-check.

## 6. Data model

A single hosted-Postgres table. Schema is created idempotently on first use.

```sql
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
);

CREATE INDEX IF NOT EXISTS ix_column_variants_report
    ON column_variants (report_key);

-- Case-insensitive uniqueness: "Management View" and "management view" collide.
CREATE UNIQUE INDEX IF NOT EXISTS ux_column_variants_report_name_ci
    ON column_variants (report_key, LOWER(name));
```

- **Name uniqueness is case-insensitive**, enforced by the `LOWER(name)` unique
  index (there is *no* plain `UNIQUE (report_key, name)`). The app's duplicate
  check is likewise case-insensitive.
- **DB-level CHECKs** back up the app validation: `report_key` must be one of the
  two known keys, and `columns` must be a non-empty JSON array. (Trade-off: adding a
  future `report_key` — e.g. for page 4 — requires `ALTER TABLE … DROP/ADD
  CONSTRAINT`; see §13.)
- **`updated_at` is not auto-maintained.** The `DEFAULT now()` fires only on
  insert; there is no trigger. Therefore **every `update_columns` and
  `rename_variant` statement must set `updated_at = now()` explicitly.**
- `report_key` constants live in `column_variants.py`.
- "Standard" is **never** stored; it is computed at runtime.
- Concurrency: **last-write-wins** (no row locking); the case-insensitive unique
  index is the only guard. Acceptable for a small internal team.

## 7. Architecture / modules

Two new modules + tests, plus thin hooks in the two pages. Engines are untouched.

- **`src/column_variants.py`** — pure data + logic layer, **no Streamlit import**:
  - Constants: `REPORT_DELIVERY_SHORTAGE`, `REPORT_SALES_ORDER_UNCONFIRMED`,
    `STANDARD_NAME = "Standard"`.
  - `@dataclass Variant(id, report_key, name, columns, created_at, updated_at)`.
  - Validation helpers: `validate_name`, `is_reserved_name`, `normalize_columns`.
  - `apply_columns(df, columns) -> df` — project & order, intersecting with present
    columns; empty/Standard → return df unchanged.
  - `VariantStore` wrapping a Postgres DSN:
    `ensure_schema()`, `list_variants(report_key)`, `get_variant(id)`,
    `create_variant(report_key, name, columns)`, `update_columns(id, columns)`,
    `rename_variant(id, name)`, `delete_variant(id)`.
    - `update_columns` and `rename_variant` **must include `updated_at = now()`** in
      their `UPDATE` statements (no DB trigger maintains it).
    - `UniqueViolation` (from the case-insensitive index) is translated into a
      friendly "a variant with that name already exists" error.
  - Each operation opens a short-lived connection via a context manager
    (robust against Neon serverless dropping idle connections). DSN is injectable
    so tests can target a throwaway database.
- **`src/column_variants_ui.py`** — Streamlit panel, **imports** `column_variants`:
  - `get_store()` cached with `@st.cache_resource` (one store/DSN per process).
  - `list_variants_cached(report_key)` via `@st.cache_data(ttl=60)`; cleared after
    any mutation so reads stay fast but never stale.
  - `render_variant_panel(report_key, all_columns) -> list[str]` renders the
    dropdown + multiselect + buttons and **returns the active ordered column list**
    for the caller to apply. Session-state keys are namespaced per `report_key`.
    State is kept minimal (selection derived from widget state; explicit refresh
    after writes) to avoid the rerun-timing bugs that plagued the old build.
- **Page hooks (small):**
  - `pages/2_Delivery_Fill_Rate_Dashboard.py`, Shortage Report tab: call the panel,
    then `display_df = apply_columns(shortage_df, cols)` before `st.dataframe`, and
    pass the projected frame as the Shortage Report sheet to `generate_excel_report`.
  - `pages/3_Sales_Order_Fill_Rate_Dashboard.py`, Unconfirmed Demand tab: same
    pattern against `view` (display) and `unconfirmed_df` (Excel).

## 8. Connection, secrets, deployment

- **Database:** a new **Neon** serverless Postgres project, free tier. Use Neon's
  **pooled** connection string (suited to Streamlit's frequent short connections);
  `sslmode=require`.
- **Driver / dependency:** `psycopg[binary]>=3.1` (psycopg 3). No ORM — parameterized
  SQL only. (One dependency; "small.")
- **Secret:** the DSN is read from `st.secrets["postgres"]["dsn"]`, falling back to
  the `DATABASE_URL` env var. Nothing is hardcoded or committed.
  - Local: `.streamlit/secrets.toml` (git-ignored).
  - Streamlit Cloud: set the same secret in the app's **Secrets** UI.
  - A committed **`.streamlit/secrets.toml.example`** documents the key.
- **`.gitignore`** (new — the repo has none): ignore `.streamlit/secrets.toml`,
  `__pycache__/`, `*.pyc`, `*.db`, `.venv/`. (Bytecode churn cleanup is a welcome
  side effect; removing already-tracked `.pyc` from the index is optional and can be
  a follow-up.)

## 9. Applying a variant (table + Excel)

- **On screen:** `apply_columns(<table_df>, cols)` then `st.dataframe(...)`.
- **Excel:** project the **same** column list onto only the variant's sheet
  (Shortage Report / Unconfirmed Demand). Other sheets (summaries, Top 10, Clean
  Data, Raw Preview) are unchanged. So "what you see is what you download" for that
  table, without disturbing the rest of the workbook.

## 10. Error handling & edge cases

- **DB unreachable / DSN missing:** the panel catches the error, shows
  "Shared variants unavailable" once, and **falls back to Standard only**. The
  dashboard never crashes because of the variants feature.
- **Saved columns absent from this upload:** intersect; render what exists. If a
  variant resolves to zero present columns, fall back to Standard with a caption.
- **Duplicate name / reserved name / empty selection:** rejected with inline errors
  (see §5 validation).
- **Variant deleted by another user mid-session:** on a missing id, fall back to
  Standard.

## 11. Testing

`tests/test_column_variants.py`:

- **Pure logic (no DB):** `validate_name`, reserved-name rejection,
  `normalize_columns`, `apply_columns` (ordering, hide, intersection with missing
  columns, empty/Standard passthrough).
- **DB integration (gated on `TEST_DATABASE_URL`, skipped if unset):** `ensure_schema`
  idempotency, create/list/get/update/rename/delete round-trips, `UNIQUE` violation
  handling, list ordering. Can target a disposable Neon branch.

## 12. Rollout / setup steps (for the implementation plan)

1. Provision the Neon project; capture the pooled DSN.
2. Add `psycopg[binary]` to `requirements.txt`; add `.gitignore` +
   `.streamlit/secrets.toml.example`; set local secret.
3. Build `column_variants.py` (+ tests) → `ensure_schema()` creates the table.
4. Build `column_variants_ui.py`.
5. Wire pages 2 & 3 (display + Excel).
6. Set the secret in Streamlit Cloud; push; verify on the deployed app.

## 13. Open questions / future

- A single shared **default** variant per report (one extra nullable flag) — omitted
  for now per scope.
- Optional: extend to the Daily Short Report (page 4) using the same mechanism with
  new report keys — this would require widening the `report_key` CHECK constraint
  (`ALTER TABLE … DROP CONSTRAINT … ADD CONSTRAINT …`).
- Optional: prettier column display labels (currently raw engine column names).
