"""Aggregate rule hits into reviewable issue groups.

Each rule result becomes one issue group with a stable ``Issue Group ID`` that
links back to all the original SAP rows involved. ``analyze()`` runs the rules,
de-duplicates identical hits, builds the one-row-per-group flagged table, the
group->row-indices map, the related-records table (for the Excel export), and
the multi-batch analysis summary.

``build_results_table()`` merges the issue groups with the automatic AI review
output and the human-review findings into the single table shown on the Flagged
Issues tab and written to the ``Flagged Issues`` Excel sheet.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from .ai_agent import AI_COLUMNS, ai_fields
from .loader import (
    COL_BATCH,
    COL_BATCH_SLED,
    COL_MATERIAL,
    COL_MATERIAL_DESC,
    COL_PLANT,
    COL_PO,
    COL_PURCH_GROUP,
    COL_RECEIVED,
    COL_SUPPLIER,
    COL_SUPPLIER_NAME,
    COL_VENDOR,
    COL_VENDOR_NAME,
    NORMALIZED_BATCH,
)
from .reviews import HUMAN_REVIEW_FIELDS
from .rules import (
    MULTI_BATCH_COLUMNS,
    RISK_ORDER,
    RuleHit,
    build_multiple_batches,
    run_all_rules,
)

# One row per issue group. Superset of the Flagged-Issues tab columns and the
# spec's issue-group fields. "Rule Risk Level" distinguishes the rule-assigned
# risk from the AI Review Priority added later.
GROUP_COLUMNS = [
    "Issue Group ID", "Issue Type", "Rule Risk Level",
    "Material", "Material Description", "Supplier", "Supplier Name",
    "Vendor", "Vendor Name", "Purchasing Group",
    "Original Batch Values", "Normalized Batch Values", "Expiry Dates",
    "Purchase Orders", "Plants", "Received Dates",
    "Number of Records", "Reason Flagged",
]


@dataclass
class AnalysisResult:
    df: pd.DataFrame                       # the full normalized receiving data
    flagged: pd.DataFrame                  # one row per issue group (GROUP_COLUMNS)
    members: dict                          # Issue Group ID -> list of df indices
    multi_batch: pd.DataFrame              # multi-batch analysis table
    summary: dict = field(default_factory=dict)


def _join_unique(series: pd.Series) -> str:
    vals = [
        str(v) for v in series
        if str(v).strip() != "" and str(v).strip().lower() not in ("nan", "nat")
    ]
    return ", ".join(dict.fromkeys(vals))


def _join_dates(series: pd.Series) -> str:
    out: list[str] = []
    for v in series:
        if pd.isna(v):
            continue
        out.append(pd.Timestamp(v).strftime("%Y-%m-%d"))
    return ", ".join(dict.fromkeys(out))


def _group_record(group_id: str, hit: RuleHit, df: pd.DataFrame) -> dict:
    g = df.loc[hit.indices]

    def col(name):
        return g[name] if name in g.columns else pd.Series([], dtype=object)

    return {
        "Issue Group ID": group_id,
        "Issue Type": hit.issue_type,
        "Rule Risk Level": hit.risk_level,
        "Material": _join_unique(col(COL_MATERIAL)),
        "Material Description": _join_unique(col(COL_MATERIAL_DESC)),
        "Supplier": _join_unique(col(COL_SUPPLIER)),
        "Supplier Name": _join_unique(col(COL_SUPPLIER_NAME)),
        "Vendor": _join_unique(col(COL_VENDOR)),
        "Vendor Name": _join_unique(col(COL_VENDOR_NAME)),
        "Purchasing Group": _join_unique(col(COL_PURCH_GROUP)),
        "Original Batch Values": _join_unique(col(COL_BATCH)),
        "Normalized Batch Values": _join_unique(col(NORMALIZED_BATCH)),
        "Expiry Dates": _join_dates(col(COL_BATCH_SLED)),
        "Purchase Orders": _join_unique(col(COL_PO)),
        "Plants": _join_unique(col(COL_PLANT)),
        "Received Dates": _join_dates(col(COL_RECEIVED)),
        "Number of Records": len(g),
        "Reason Flagged": hit.reason,
    }


def _dedup_hits(hits: list[RuleHit]) -> list[RuleHit]:
    """Drop hits that flag the same records for the same issue type (some rules
    can re-derive an identical group). First occurrence wins."""
    seen: set = set()
    out: list[RuleHit] = []
    for h in hits:
        key = (h.issue_type, frozenset(h.indices))
        if key in seen:
            continue
        seen.add(key)
        out.append(h)
    return out


def analyze(df: pd.DataFrame) -> AnalysisResult:
    """Run all rules and assemble issue groups, sorted High→Low risk then by
    record count (biggest first). One issue group per unique rule result."""
    hits = _dedup_hits(run_all_rules(df))
    hits.sort(key=lambda h: (RISK_ORDER.get(h.risk_level, 9), -len(h.indices)))

    records: list[dict] = []
    members: dict = {}
    for i, hit in enumerate(hits, 1):
        gid = f"IG-{i:04d}"
        members[gid] = list(hit.indices)
        records.append(_group_record(gid, hit, df))

    flagged = pd.DataFrame(records, columns=GROUP_COLUMNS)
    multi_batch = build_multiple_batches(df)

    risk_counts = flagged["Rule Risk Level"].value_counts().to_dict() if not flagged.empty else {}
    summary = {
        "total_records": int(len(df)),
        "unique_materials": int(df[COL_MATERIAL].replace("", pd.NA).nunique()),
        "unique_batches": int(df[COL_BATCH].replace("", pd.NA).nunique()),
        "unique_suppliers": int(df[COL_SUPPLIER].replace("", pd.NA).nunique()) if COL_SUPPLIER in df.columns else 0,
        "unique_pos": int(df[COL_PO].replace("", pd.NA).nunique()) if COL_PO in df.columns else 0,
        "total_issue_groups": int(len(flagged)),
        "high_risk": int(risk_counts.get("High", 0)),
        "medium_risk": int(risk_counts.get("Medium", 0)),
        "low_risk": int(risk_counts.get("Low", 0)),
        "multi_batch_materials": int(len(multi_batch)),
    }
    return AnalysisResult(df=df, flagged=flagged, members=members,
                          multi_batch=multi_batch, summary=summary)


def build_related_records(result: AnalysisResult) -> pd.DataFrame:
    """All original SAP rows behind every flagged issue group, each prefixed
    with its Issue Group ID / Issue Type / Rule Risk Level / Reason Flagged.
    Internal helper columns are dropped."""
    src_cols = [c for c in result.df.columns if c != NORMALIZED_BATCH]
    blocks: list[pd.DataFrame] = []
    for _, row in result.flagged.iterrows():
        gid = row["Issue Group ID"]
        idx = result.members.get(gid, [])
        if not idx:
            continue
        block = result.df.loc[idx, src_cols].copy()
        block.insert(0, "Reason Flagged", row["Reason Flagged"])
        block.insert(0, "Rule Risk Level", row["Rule Risk Level"])
        block.insert(0, "Issue Type", row["Issue Type"])
        block.insert(0, "Issue Group ID", gid)
        blocks.append(block)
    if not blocks:
        cols = ["Issue Group ID", "Issue Type", "Rule Risk Level", "Reason Flagged"] + src_cols
        return pd.DataFrame(columns=cols)
    return pd.concat(blocks, ignore_index=True)


# Columns of the merged results table, in display/export order.
RESULTS_COLUMNS = GROUP_COLUMNS + AI_COLUMNS + HUMAN_REVIEW_FIELDS


def build_results_table(
    flagged: pd.DataFrame,
    ai_map: dict | None = None,
    review_store: dict | None = None,
    *,
    unavailable_label: str = "Not Available",
) -> pd.DataFrame:
    """Merge issue groups with the automatic AI review output and the human
    findings into one table. ``ai_map`` maps Issue Group ID -> AIReviewResult
    (or ``None``/``"skipped"``). ``review_store`` maps Issue Group ID -> human
    review dict."""
    ai_map = ai_map or {}
    review_store = review_store or {}
    if flagged.empty:
        return pd.DataFrame(columns=RESULTS_COLUMNS)
    rows: list[dict] = []
    for _, r in flagged.iterrows():
        gid = r["Issue Group ID"]
        row = r.to_dict()
        row.update(ai_fields(ai_map.get(gid), unavailable_label=unavailable_label))
        review = review_store.get(gid, {})
        row.update({k: review.get(k, "") for k in HUMAN_REVIEW_FIELDS})
        rows.append(row)
    return pd.DataFrame(rows, columns=RESULTS_COLUMNS)
