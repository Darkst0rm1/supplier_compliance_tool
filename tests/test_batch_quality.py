"""Tests for the Batch Quality Analysis engine."""
from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from src.batch_quality import ai_review, exporter, loader, reviews
from src.batch_quality.issue_groups import analyze, build_related_records
from src.batch_quality.loader import (
    COL_BATCH,
    COL_BATCH_SLED,
    COL_MATERIAL,
    COL_PLANT,
    COL_PO,
    COL_QTY,
    COL_RECEIVED,
    COL_SUPPLIER,
    COL_SUPPLIER_NAME,
    COL_VENDOR_NAME,
    NORMALIZED_BATCH,
)
from src.batch_quality import rules


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _xlsx(headers: list[str], rows: list[list]) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = loader.SAP_SHEET
    ws.append(headers)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


_CANON_COLS = [
    COL_MATERIAL, "Material Desc", COL_BATCH, COL_BATCH_SLED, COL_PO,
    COL_PLANT, COL_SUPPLIER, COL_SUPPLIER_NAME, COL_VENDOR_NAME,
    COL_RECEIVED, COL_QTY,
]


def make_df(records: list[dict]) -> pd.DataFrame:
    """Build a DataFrame shaped like loader output (canonical cols + dates +
    Normalized Batch) from concise record dicts."""
    rows = []
    for i, rec in enumerate(records):
        rows.append({
            COL_MATERIAL: str(rec.get("material", "1000")),
            "Material Desc": rec.get("desc", "WIDGET"),
            COL_BATCH: rec.get("batch", "B1"),
            COL_BATCH_SLED: rec.get("sled"),
            COL_PO: str(rec.get("po", f"PO{i}")),
            COL_PLANT: str(rec.get("plant", "2910")),
            COL_SUPPLIER: str(rec.get("supplier", "S1")),
            COL_SUPPLIER_NAME: rec.get("supplier_name", "Supplier One"),
            COL_VENDOR_NAME: rec.get("vendor_name", "Vendor One"),
            COL_RECEIVED: rec.get("received"),
            COL_QTY: rec.get("qty", 1.0),
        })
    df = pd.DataFrame(rows, columns=_CANON_COLS)
    df[COL_BATCH_SLED] = pd.to_datetime(df[COL_BATCH_SLED], errors="coerce")
    df[COL_RECEIVED] = pd.to_datetime(df[COL_RECEIVED], errors="coerce")
    df[NORMALIZED_BATCH] = df[COL_BATCH].map(loader.normalize_batch)
    return df.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Loader / normalization
# ---------------------------------------------------------------------------
def test_normalize_batch_variants():
    assert loader.normalize_batch("31-5357") == "315357"
    assert loader.normalize_batch("31 5357") == "315357"
    assert loader.normalize_batch("31/5357") == "315357"
    assert loader.normalize_batch("31.53_57") == "315357"
    assert loader.normalize_batch("  abc1  ") == "ABC1"
    assert loader.normalize_batch(None) == ""
    assert loader.normalize_batch(float("nan")) == ""


def test_load_normalizes_headers_and_parses():
    # Alias headers ("Purchase Order", "Material Description") + Material as float
    # (.0 artifact), Batch as text with leading zeros, a real datetime SLED.
    buf = _xlsx(
        ["Material", "Material Description", "Batch", "Batch SLED",
         "Purchase Order", "Plant", "Qty Base"],
        [
            [10038128.0, "SCHAR", "00009926", datetime(2027, 1, 9), 1000006096, 2910, 675],
            [10038095.0, "HONEY", "310327A", datetime(2027, 3, 31), 1000006096, 2910, 507],
        ],
    )
    df = loader.load_batch_data(buf)
    assert COL_PO in df.columns and "Material Desc" in df.columns
    # Text id: leading zeros preserved, trailing .0 stripped
    assert df[COL_BATCH].iloc[0] == "00009926"
    assert df[COL_MATERIAL].iloc[0] == "10038128"
    # Dates parsed
    assert pd.api.types.is_datetime64_any_dtype(df[COL_BATCH_SLED])
    assert df[COL_BATCH_SLED].iloc[0] == pd.Timestamp(2027, 1, 9)
    # Normalized batch present
    assert df[NORMALIZED_BATCH].iloc[1] == "310327A"


def test_load_missing_columns_raises():
    buf = _xlsx(["Foo", "Bar"], [[1, 2]])
    with pytest.raises(loader.BatchQualityError):
        loader.load_batch_data(buf)


# ---------------------------------------------------------------------------
# Rules
# ---------------------------------------------------------------------------
def test_rule_format_variation():
    df = make_df([
        {"material": "M1", "batch": "31-5357", "sled": "2027-01-09", "po": "PO1"},
        {"material": "M1", "batch": "315357", "sled": "2027-01-09", "po": "PO2"},
    ])
    hits = rules.rule_batch_groups(df)
    assert len(hits) == 1
    assert hits[0].issue_type == rules.ISSUE_FORMAT_VARIATION
    assert hits[0].risk_level == rules.RISK_MEDIUM
    assert set(hits[0].indices) == {0, 1}


def test_rule_conflicting_expiry_takes_priority():
    df = make_df([
        {"material": "M1", "batch": "ABC1", "sled": "2027-01-09", "po": "PO1"},
        {"material": "M1", "batch": "ABC1", "sled": "2027-02-09", "po": "PO2"},
    ])
    hits = rules.rule_batch_groups(df)
    assert len(hits) == 1
    assert hits[0].issue_type == rules.ISSUE_CONFLICTING_EXPIRY
    assert hits[0].risk_level == rules.RISK_HIGH


def test_rule_duplicate_receiving():
    df = make_df([
        {"material": "M1", "batch": "B1", "sled": "2027-01-09", "po": "PO9", "qty": 5.0},
        {"material": "M1", "batch": "B1", "sled": "2027-01-09", "po": "PO9", "qty": 7.0},
    ])
    hits = rules.rule_duplicate_receiving(df)
    assert len(hits) == 1
    assert hits[0].issue_type == rules.ISSUE_DUPLICATE_RECEIVING


def test_rule_exact_duplicate():
    df = make_df([
        {"material": "M1", "batch": "B1", "sled": "2027-01-09", "po": "PO9", "qty": 5.0},
        {"material": "M1", "batch": "B1", "sled": "2027-01-09", "po": "PO9", "qty": 5.0},
    ])
    hits = rules.rule_exact_duplicate(df)
    assert len(hits) == 1
    assert hits[0].issue_type == rules.ISSUE_EXACT_DUPLICATE
    assert hits[0].risk_level == rules.RISK_HIGH


def test_rule_format_review():
    df = make_df([
        {"material": "M1", "batch": "B1 ", "sled": "2027-01-09", "po": "PO1"},   # trailing space
        {"material": "M2", "batch": "X", "sled": "2027-01-09", "po": "PO2"},      # too short
        {"material": "M3", "batch": "GOODBATCH", "sled": "2027-01-09", "po": "PO3"},
    ])
    hits = rules.rule_format_review(df)
    types = {h.issue_type for h in hits}
    assert types == {rules.ISSUE_FORMAT_REVIEW}
    assert len(hits) == 2  # only the two malformed batches


def test_multiple_batches_summary():
    df = make_df([
        {"material": "M1", "batch": "B1", "sled": "2027-01-01"},
        {"material": "M1", "batch": "B2", "sled": "2027-02-01"},  # multi batch + multi date
        {"material": "M2", "batch": "B3", "sled": "2027-01-01"},
        {"material": "M2", "batch": "B4", "sled": "2027-01-01"},  # multi batch, single date
    ])
    mb = rules.build_multiple_batches(df)
    assert list(mb["Material"]) == ["M1"]
    assert mb.iloc[0]["Number of Batch Values"] == 2
    assert mb.iloc[0]["Number of Expiry Dates"] == 2


# ---------------------------------------------------------------------------
# Issue groups
# ---------------------------------------------------------------------------
def test_analyze_issue_groups_and_members():
    df = make_df([
        {"material": "M1", "batch": "ABC1", "sled": "2027-01-09", "po": "PO1"},
        {"material": "M1", "batch": "ABC1", "sled": "2027-02-09", "po": "PO2"},
        {"material": "M2", "batch": "31-5357", "sled": "2027-03-01", "po": "PO3"},
        {"material": "M2", "batch": "315357", "sled": "2027-03-01", "po": "PO4"},
    ])
    result = analyze(df)
    assert result.summary["total_issue_groups"] == 2
    # High risk sorts first
    assert result.flagged.iloc[0]["Risk Level"] == "High"
    gid = result.flagged.iloc[0]["Issue Group ID"]
    assert gid == "IG-0001"
    assert set(result.members[gid]) == {0, 1}


def test_build_related_records_links_rows():
    df = make_df([
        {"material": "M1", "batch": "ABC1", "sled": "2027-01-09", "po": "PO1"},
        {"material": "M1", "batch": "ABC1", "sled": "2027-02-09", "po": "PO2"},
    ])
    result = analyze(df)
    related = build_related_records(result)
    assert "Issue Group ID" in related.columns
    assert len(related) == 2
    assert NORMALIZED_BATCH not in related.columns


# ---------------------------------------------------------------------------
# AI review
# ---------------------------------------------------------------------------
def test_ai_context_only_relevant_records():
    df = make_df([
        {"material": "A", "batch": "ABC1", "sled": "2027-01-09", "po": "PO1"},
        {"material": "A", "batch": "ABC1", "sled": "2027-02-09", "po": "PO2"},
        {"material": "B", "batch": "ZZZ9", "sled": "2027-05-01", "po": "PO3"},
        {"material": "C", "batch": "QQQ8", "sled": "2027-06-01", "po": "PO4"},
    ])
    result = analyze(df)
    gid = result.flagged.iloc[0]["Issue Group ID"]
    idx = result.members[gid]
    related = df.loc[idx]
    mat = related[COL_MATERIAL].iloc[0]
    material_records = df[df[COL_MATERIAL] == mat]
    nbs = list(related[NORMALIZED_BATCH].unique())
    normbatch_records = df[df[NORMALIZED_BATCH].isin(nbs)]

    grp_row = result.flagged.iloc[0].to_dict()
    ctx = ai_review.build_ai_context(grp_row, related, material_records, normbatch_records)

    assert len(ctx["related_records"]) == 2
    # Only material A is in scope — unrelated materials B and C never appear.
    for bucket in ("related_records", "material_records", "normalized_batch_records"):
        for rec in ctx[bucket]:
            assert rec["Material"] == "A"


def test_ai_schema_has_no_corrected_value_fields():
    fields = set(ai_review.AIReviewResult.model_fields.keys())
    assert fields.isdisjoint(ai_review.FORBIDDEN_FIELDS)
    # No field name hints at a corrected/suggested/replacement value
    for f in fields:
        assert "correct" not in f and "suggest" not in f and "replacement" not in f


def test_get_model_default(monkeypatch):
    monkeypatch.delenv("ANTHROPIC_MODEL", raising=False)
    assert ai_review.get_model() == ai_review.DEFAULT_MODEL
    assert ai_review.get_model("claude-sonnet-4-6") == "claude-sonnet-4-6"


def test_is_ai_configured(monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    assert ai_review.is_ai_configured() is False
    assert ai_review.is_ai_configured("sk-test") is True


# ---------------------------------------------------------------------------
# Human reviews
# ---------------------------------------------------------------------------
def test_save_and_get_review():
    store: dict = {}
    reviews.save_review(store, "IG-0001", {
        "Confirmed Issue": "Yes",
        "Root Cause": "Receiving data entry",
        "Reviewer Comment": "Looks like a typo.",
        "Ignored": "dropped",
    })
    saved = reviews.get_review(store, "IG-0001")
    assert saved["Confirmed Issue"] == "Yes"
    assert saved["Root Cause"] == "Receiving data entry"
    assert "Ignored" not in saved
    # Unknown id returns blank template
    blank = reviews.get_review(store, "IG-9999")
    assert blank["Confirmed Issue"] == ""


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------
def test_generate_excel_has_four_sheets():
    df = make_df([
        {"material": "M1", "batch": "ABC1", "sled": "2027-01-09", "po": "PO1"},
        {"material": "M1", "batch": "ABC1", "sled": "2027-02-09", "po": "PO2"},
    ])
    result = analyze(df)
    related = build_related_records(result)
    xlsx = exporter.generate_excel(result.flagged, related, result.multi_batch, pd.DataFrame())
    wb = load_workbook(io.BytesIO(xlsx))
    assert wb.sheetnames == ["Flagged Issues", "Related Records", "Multiple Batches", "AI Review"]


# ---------------------------------------------------------------------------
# App wiring
# ---------------------------------------------------------------------------
def test_app_registers_batch_quality_page():
    with open("app.py", encoding="utf-8") as f:
        content = f.read()
    assert "8_Batch_Quality_Analysis.py" in content
