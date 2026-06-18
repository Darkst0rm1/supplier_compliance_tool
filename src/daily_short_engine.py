"""Processing engine for the SAPUI5 Daily Short Report export.

The export is order-line level. Each line carries an ordered quantity and the
quantities that made it through each fulfilment stage:

    Order Quantity -> Confirmed -> Total Delivery -> Picked -> Invoice

The workbook embeds a TEMPLATE at the bottom (below the data) describing the
report. This engine implements that template:

* Summary totals + fill-rate ratios (Confirmed/Order, Delivery/Order, Invoice/Order)
* Three "shorted" analyses, each Order-based:
    - Unconfirmed quantities            = Order - Confirmed
    - Shorted at outbound delivery       = Order - Delivered
    - Shorted at invoicing               = Order - Invoiced
* Each analysis is a table with the template columns:
    Plant | Sales order # | Customer | Material | Material description |
    Ordered | Confirmed | Shorted | Reason
"""
from __future__ import annotations

import io
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows


# ---------------------------------------------------------------------------
# Source columns (exact export headers) + tolerant aliases
# ---------------------------------------------------------------------------
COL_SALES_ORDER   = "Sales Order"
COL_MATERIAL      = "Material"
COL_ORDER_QTY     = "Order Quantity"
COL_CONFIRMED_QTY = "Confirmed Quantity (CS)"
COL_DELIVERY_QTY  = "Total Delivery (Sales) Quantity (CS)"
COL_PICKED_QTY    = "Picked Quantity (CS)"
COL_INVOICE_QTY   = "Invoice Quantity (CS)"
COL_PLANT         = "Plant"
COL_CUSTOMER      = "Sold To Name"
COL_MAT_DESC      = "TOL Material Description"
COL_REASON        = "Reason for Rejection Desc."
COL_ODS           = "Outbound Delivery Status"
COL_ORDER_TYPE    = "Sales Order Type Desc."
COL_NET_AMT       = "Item Net Amount"
COL_NET_AMT_CONF  = "Item Net Amount (Confirmed)"
COL_VENDOR_NAME   = "Vendor Name"
COL_CDM           = "CDM Name"
COL_REQ_DATE      = "Requested Delivery Date"
COL_SHIP_DATE     = "Shipped Date"

_QTY_COLS = [COL_ORDER_QTY, COL_CONFIRMED_QTY, COL_DELIVERY_QTY, COL_PICKED_QTY, COL_INVOICE_QTY]

# Header aliases -> canonical, applied case-insensitively after stripping.
_ALIASES = {
    "sold to name": COL_CUSTOMER,
    "customer": COL_CUSTOMER,
    "tol material description": COL_MAT_DESC,
    "material description": COL_MAT_DESC,
    "total delivery (sales) quantity (cs)": COL_DELIVERY_QTY,
    "delivered quantity (cs)": COL_DELIVERY_QTY,
    "reason for rejection desc.": COL_REASON,
    "reason for rejection": COL_REASON,
}


# ---------------------------------------------------------------------------
# Short analyses (template) — Order-based
# ---------------------------------------------------------------------------
SHORT_DEFS: list[dict[str, str]] = [
    {"key": "unconfirmed", "label": "Unconfirmed Quantities",
     "title": "Shorted at time of confirmation", "qty_col": COL_CONFIRMED_QTY},
    {"key": "delivery", "label": "Shorted at Outbound Delivery",
     "title": "Shorted at time of outbound delivery", "qty_col": COL_DELIVERY_QTY},
    {"key": "invoice", "label": "Shorted at Invoicing",
     "title": "Shorted at time of invoicing", "qty_col": COL_INVOICE_QTY},
]


def _short_col(key: str) -> str:
    return f"_short_{key}"


# ---------------------------------------------------------------------------
# Loader
# ---------------------------------------------------------------------------
def _normalize_header(name: Any) -> str:
    return str(name).strip()


def _apply_aliases(df: pd.DataFrame) -> pd.DataFrame:
    rename = {}
    for col in df.columns:
        key = str(col).strip().lower()
        if key in _ALIASES and _ALIASES[key] != col:
            rename[col] = _ALIASES[key]
    return df.rename(columns=rename) if rename else df


def load_daily_short(file_obj: Any) -> pd.DataFrame:
    """Read the export, drop the embedded template/footer rows, coerce types,
    and compute the three Order-based short columns. Returns the clean lines."""
    file_obj.seek(0)
    df = pd.read_excel(file_obj, engine="openpyxl")
    df.columns = [_normalize_header(c) for c in df.columns]
    df = _apply_aliases(df)

    missing = [c for c in [COL_SALES_ORDER, COL_ORDER_QTY] if c not in df.columns]
    if missing:
        raise DailyShortError(
            "This doesn't look like a Daily Short export — missing column(s): "
            + ", ".join(missing)
        )

    # Coerce quantity columns to numeric (footer/template rows carry text and
    # become NaN, which lets us drop them cleanly).
    for c in _QTY_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Keep only real order lines: a Sales Order present and a numeric order qty.
    df = df[df[COL_SALES_ORDER].notna() & df[COL_ORDER_QTY].notna()].copy()
    df = df.reset_index(drop=True)

    # Normalise text id columns to clean strings.
    for c in [COL_SALES_ORDER, COL_MATERIAL, COL_PLANT]:
        if c in df.columns:
            df[c] = (
                df[c].astype(str)
                .str.replace(r"\.0$", "", regex=True)
                .str.strip()
            )

    # Dates
    for c in [COL_REQ_DATE, COL_SHIP_DATE]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Short columns (clip at 0 — never count overdelivery as negative short).
    order = df[COL_ORDER_QTY].fillna(0)
    for d in SHORT_DEFS:
        qcol = d["qty_col"]
        fulfilled = df[qcol].fillna(0) if qcol in df.columns else 0
        df[_short_col(d["key"])] = (order - fulfilled).clip(lower=0)

    return df


class DailyShortError(Exception):
    """Raised when the uploaded file isn't a usable Daily Short export."""


# ---------------------------------------------------------------------------
# Summary / KPIs
# ---------------------------------------------------------------------------
def build_kpis(df: pd.DataFrame) -> dict[str, Any]:
    def _sum(col: str) -> float:
        return float(df[col].sum()) if col in df.columns else 0.0

    ordered   = _sum(COL_ORDER_QTY)
    confirmed = _sum(COL_CONFIRMED_QTY)
    delivered = _sum(COL_DELIVERY_QTY)
    invoiced  = _sum(COL_INVOICE_QTY)

    def _rate(num: float, den: float) -> float:
        return (num / den * 100) if den else 0.0

    return {
        "lines": int(len(df)),
        "ordered": ordered,
        "confirmed": confirmed,
        "delivered": delivered,
        "invoiced": invoiced,
        "confirm_rate": _rate(confirmed, ordered),
        "delivery_rate": _rate(delivered, ordered),
        "invoice_rate": _rate(invoiced, ordered),
        "delivery_vs_confirmed": _rate(delivered, confirmed),
        "invoice_vs_delivered": _rate(invoiced, delivered),
        "short_unconfirmed": float(df[_short_col("unconfirmed")].sum()) if _short_col("unconfirmed") in df else 0.0,
        "short_delivery": float(df[_short_col("delivery")].sum()) if _short_col("delivery") in df else 0.0,
        "short_invoice": float(df[_short_col("invoice")].sum()) if _short_col("invoice") in df else 0.0,
        "lines_unconfirmed": int((df[_short_col("unconfirmed")] > 0).sum()) if _short_col("unconfirmed") in df else 0,
        "lines_delivery": int((df[_short_col("delivery")] > 0).sum()) if _short_col("delivery") in df else 0,
        "lines_invoice": int((df[_short_col("invoice")] > 0).sum()) if _short_col("invoice") in df else 0,
    }


# ---------------------------------------------------------------------------
# Short tables (template columns)
# ---------------------------------------------------------------------------
# Output column order exactly as the embedded template specifies.
TEMPLATE_COLUMNS = [
    "Plant", "Sales order #", "Customer", "Material", "Material description",
    "Ordered", "Confirmed", "Shorted", "Reason",
]


def build_short_table(df: pd.DataFrame, key: str, top_n: int | None = None) -> pd.DataFrame:
    """Build the template table for one short analysis: only lines with a short,
    sorted by largest short. ``top_n=None`` returns all."""
    scol = _short_col(key)
    if scol not in df.columns:
        return pd.DataFrame(columns=TEMPLATE_COLUMNS)

    sub = df[df[scol] > 0].copy()
    sub = sub.sort_values(scol, ascending=False)
    if top_n is not None:
        sub = sub.head(top_n)

    out = pd.DataFrame({
        "Plant":               sub.get(COL_PLANT),
        "Sales order #":       sub.get(COL_SALES_ORDER),
        "Customer":            sub.get(COL_CUSTOMER),
        "Material":            sub.get(COL_MATERIAL),
        "Material description": sub.get(COL_MAT_DESC),
        "Ordered":             sub.get(COL_ORDER_QTY),
        "Confirmed":           sub.get(COL_CONFIRMED_QTY),
        "Shorted":             sub[scol],
        "Reason":              sub.get(COL_REASON),
    })
    return out.reset_index(drop=True)


def build_group_summary(df: pd.DataFrame, group_col: str, key: str) -> pd.DataFrame | None:
    """Roll a short analysis up by a grouping column (Customer, Vendor, …)."""
    scol = _short_col(key)
    if group_col not in df.columns or scol not in df.columns:
        return None
    g = (
        df.groupby(group_col)
        .agg(Ordered=(COL_ORDER_QTY, "sum"), Shorted=(scol, "sum"),
             Lines=(scol, lambda s: int((s > 0).sum())))
        .reset_index()
    )
    g = g[g["Shorted"] > 0].sort_values("Shorted", ascending=False)
    g.rename(columns={group_col: group_col}, inplace=True)
    return g.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Excel report
# ---------------------------------------------------------------------------
_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_KPI_FILL    = PatternFill("solid", fgColor="2E75B6")
_ALT_FILL    = PatternFill("solid", fgColor="D6E4F0")
_HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
_TITLE_FONT  = Font(bold=True, color="1F4E79", size=14)
_KPI_LABEL   = Font(bold=True, color="FFFFFF", size=10)
_KPI_VALUE   = Font(bold=True, color="FFFFFF", size=12)
_THIN        = Side(style="thin", color="BFBFBF")
_BORDER      = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER      = Alignment(horizontal="center", vertical="center")
_LEFT        = Alignment(horizontal="left", vertical="center")


def _autofit(ws, min_w=8, max_w=48):
    for col_cells in ws.columns:
        length = min_w
        letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value is not None:
                length = max(length, min(max_w, len(str(cell.value)) + 2))
        ws.column_dimensions[letter].width = length


def _write_df(ws, df: pd.DataFrame, start_row: int = 1) -> int:
    for c_idx, col in enumerate(df.columns, 1):
        cell = ws.cell(row=start_row, column=c_idx, value=col)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = _CENTER
        cell.border = _BORDER
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start_row + 1):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = _BORDER
            cell.alignment = _LEFT
            if r_idx % 2 == 0:
                cell.fill = _ALT_FILL
    return start_row + len(df) + 2


def _summary_sheet(ws, kpis: dict[str, Any]):
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:C1")
    t = ws.cell(row=1, column=1, value="Daily Short Report — Summary")
    t.font = _TITLE_FONT
    t.alignment = _CENTER
    ws.row_dimensions[1].height = 28

    rows = [
        ("Order Lines",            f"{kpis['lines']:,}"),
        ("Total Ordered",          f"{kpis['ordered']:,.0f}"),
        ("Total Confirmed",        f"{kpis['confirmed']:,.0f}"),
        ("Total Delivered",        f"{kpis['delivered']:,.0f}"),
        ("Total Invoiced",         f"{kpis['invoiced']:,.0f}"),
        ("Confirmation Rate",      f"{kpis['confirm_rate']:.2f}%"),
        ("Delivery Fill Rate",     f"{kpis['delivery_rate']:.2f}%"),
        ("Invoice Fill Rate",      f"{kpis['invoice_rate']:.2f}%"),
        ("Unconfirmed Short (qty)", f"{kpis['short_unconfirmed']:,.0f}"),
        ("Delivery Short (qty)",   f"{kpis['short_delivery']:,.0f}"),
        ("Invoice Short (qty)",    f"{kpis['short_invoice']:,.0f}"),
    ]
    r = 3
    for label, value in rows:
        lc = ws.cell(row=r, column=1, value=label)
        vc = ws.cell(row=r, column=2, value=value)
        lc.fill = vc.fill = _KPI_FILL
        lc.font = _KPI_LABEL; lc.alignment = _LEFT; lc.border = _BORDER
        vc.font = _KPI_VALUE; vc.alignment = _CENTER; vc.border = _BORDER
        ws.row_dimensions[r].height = 20
        r += 1
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18


def generate_excel_report(df: pd.DataFrame, kpis: dict[str, Any], top_n: int | None = None) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    _summary_sheet(wb.create_sheet("Summary"), kpis)

    sheet_names = {
        "unconfirmed": "Top Unconfirmed",
        "delivery": "Top Shorted at Delivery",
        "invoice": "Top Shorted at Invoicing",
    }
    for d in SHORT_DEFS:
        ws = wb.create_sheet(sheet_names[d["key"]])
        table = build_short_table(df, d["key"], top_n=top_n)
        title = d["title"] + (f" — Top {top_n}" if top_n else " — all")
        ws.cell(row=1, column=1, value=title).font = _TITLE_FONT
        if table.empty:
            ws.cell(row=2, column=1, value="No shorts found.")
        else:
            _write_df(ws, table, start_row=2)
        _autofit(ws)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
