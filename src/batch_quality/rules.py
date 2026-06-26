"""Rule-based batch-quality detectors.

Each rule scans the normalized receiving DataFrame and returns ``RuleHit``s — a
flag plus the original SAP row indices involved. No row is ever modified or
deleted; the rules only *identify* records for human review. AI is not involved
here.

Rules (from the business spec):

1. Same probable batch in different formats        -> Batch Format Variation (Medium)
2. Same probable batch, conflicting expiry dates   -> Conflicting Expiry (High)
3. Duplicate PO+Material+Batch+SLED combination    -> Duplicate Receiving (Medium)
4. Completely identical row                         -> Exact Duplicate (High)
5. Batch formatting concern (spaces/punct/short/…) -> Batch Format Review (Low)
6. (analysis only) materials with >1 batch AND >1 expiry date — not an issue flag
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

from .loader import (
    COL_BATCH,
    COL_BATCH_SLED,
    COL_MATERIAL,
    COL_MATERIAL_DESC,
    COL_PO,
    COL_RECEIVED,
    COL_SUPPLIER_NAME,
    NORMALIZED_BATCH,
)

# ---------------------------------------------------------------------------
# Issue types and risk levels
# ---------------------------------------------------------------------------
ISSUE_FORMAT_VARIATION = "Batch Format Variation"
ISSUE_CONFLICTING_EXPIRY = "Probable Same Batch With Conflicting Expiry Dates"
ISSUE_DUPLICATE_RECEIVING = "Duplicate Receiving Combination"
ISSUE_EXACT_DUPLICATE = "Exact Duplicate Record"
ISSUE_FORMAT_REVIEW = "Batch Format Review"

RISK_HIGH = "High"
RISK_MEDIUM = "Medium"
RISK_LOW = "Low"
RISK_ORDER = {RISK_HIGH: 0, RISK_MEDIUM: 1, RISK_LOW: 2}


@dataclass
class RuleHit:
    """One flagged group: an issue type/risk, a reason, and the SAP rows."""

    issue_type: str
    risk_level: str
    reason: str
    indices: list = field(default_factory=list)


def _nonblank(series: pd.Series) -> pd.Series:
    return series[series.astype(str).str.strip() != ""]


# ---------------------------------------------------------------------------
# Rule 1 + 2: same normalized batch, grouped by Material
# ---------------------------------------------------------------------------
def rule_batch_groups(df: pd.DataFrame) -> list[RuleHit]:
    """Group by Material + Normalized Batch. If the group has more than one
    distinct nonblank Batch SLED -> conflicting expiry (High, Rule 2). Else if
    it has more than one original Batch format -> format variation (Medium,
    Rule 1)."""
    hits: list[RuleHit] = []
    has_sled = COL_BATCH_SLED in df.columns
    sub = df[df[COL_BATCH].astype(str).str.strip() != ""]
    for (material, nb), g in sub.groupby([COL_MATERIAL, NORMALIZED_BATCH]):
        if str(nb).strip() == "":
            continue
        formats = sorted(set(g[COL_BATCH].astype(str)))
        distinct_dates = (
            _nonblank(g[COL_BATCH_SLED]).dropna().unique() if has_sled else []
        )
        if len(distinct_dates) > 1:
            dates = sorted(pd.Timestamp(d).strftime("%Y-%m-%d") for d in distinct_dates)
            hits.append(RuleHit(
                ISSUE_CONFLICTING_EXPIRY, RISK_HIGH,
                f"Material {material}: normalized batch {nb} appears with "
                f"conflicting expiry dates {dates} across formats {formats}.",
                list(g.index),
            ))
        elif len(formats) > 1:
            hits.append(RuleHit(
                ISSUE_FORMAT_VARIATION, RISK_MEDIUM,
                f"Material {material}: normalized batch {nb} entered in "
                f"{len(formats)} different formats {formats}.",
                list(g.index),
            ))
    return hits


# ---------------------------------------------------------------------------
# Rule 3: duplicate receiving combination
# ---------------------------------------------------------------------------
def rule_duplicate_receiving(df: pd.DataFrame) -> list[RuleHit]:
    keys = [c for c in (COL_PO, COL_MATERIAL, COL_BATCH, COL_BATCH_SLED) if c in df.columns]
    if COL_MATERIAL not in keys or COL_BATCH not in keys:
        return []
    hits: list[RuleHit] = []
    for combo, g in df.groupby(keys, dropna=False):
        if len(g) <= 1:
            continue
        if str(g[COL_BATCH].iloc[0]).strip() == "":
            continue
        combo = combo if isinstance(combo, tuple) else (combo,)
        label = ", ".join(f"{k}={v}" for k, v in zip(keys, combo))
        hits.append(RuleHit(
            ISSUE_DUPLICATE_RECEIVING, RISK_MEDIUM,
            f"{len(g)} receiving records share the same combination ({label}). "
            "May be a valid split receipt — review before any action.",
            list(g.index),
        ))
    return hits


# ---------------------------------------------------------------------------
# Rule 4: exact duplicate row
# ---------------------------------------------------------------------------
def rule_exact_duplicate(df: pd.DataFrame) -> list[RuleHit]:
    src_cols = [c for c in df.columns if c != NORMALIZED_BATCH]
    dup_mask = df.duplicated(subset=src_cols, keep=False)
    if not dup_mask.any():
        return []
    hits: list[RuleHit] = []
    for _, g in df[dup_mask].groupby(src_cols, dropna=False):
        if len(g) <= 1:
            continue
        hits.append(RuleHit(
            ISSUE_EXACT_DUPLICATE, RISK_HIGH,
            f"{len(g)} completely identical receiving rows.",
            list(g.index),
        ))
    return hits


# ---------------------------------------------------------------------------
# Rule 5: batch formatting concern (per distinct Material+Batch)
# ---------------------------------------------------------------------------
def _format_concerns(batch: object) -> list[str]:
    if batch is None:
        return []
    s = str(batch)
    if s.strip() == "":
        return []
    concerns: list[str] = []
    if s != s.strip():
        concerns.append("leading or trailing spaces")
    if re.search(r"\S\s{2,}\S", s):
        concerns.append("repeated internal spaces")
    core = s.strip()
    if re.search(r"[^\w]$", core):
        concerns.append("unexpected trailing punctuation")
    if len(core) <= 2:
        concerns.append("only one or two characters")
    elif len(set(core)) == 1:
        concerns.append("one character repeated for the whole batch")
    return concerns


def rule_format_review(df: pd.DataFrame) -> list[RuleHit]:
    flagged: dict[tuple, tuple[list, list[str]]] = {}
    for idx, batch, material in zip(df.index, df[COL_BATCH], df[COL_MATERIAL]):
        concerns = _format_concerns(batch)
        if not concerns:
            continue
        key = (material, str(batch))
        flagged.setdefault(key, ([], concerns))[0].append(idx)
    hits: list[RuleHit] = []
    for (material, batch), (indices, concerns) in flagged.items():
        hits.append(RuleHit(
            ISSUE_FORMAT_REVIEW, RISK_LOW,
            f"Material {material}, batch {batch!r}: {', '.join(concerns)}. "
            "Requires review but is not automatically incorrect.",
            indices,
        ))
    return hits


def run_all_rules(df: pd.DataFrame) -> list[RuleHit]:
    """Run every rule (1–5) and return all hits."""
    return (
        rule_batch_groups(df)
        + rule_duplicate_receiving(df)
        + rule_exact_duplicate(df)
        + rule_format_review(df)
    )


# ---------------------------------------------------------------------------
# Rule 6 (analysis only): materials with multiple batches AND multiple dates
# ---------------------------------------------------------------------------
MULTI_BATCH_COLUMNS = [
    "Material", "Material Description", "Supplier Name",
    "Number of Batch Values", "Number of Expiry Dates",
    "Earliest Expiry Date", "Latest Expiry Date",
    "Number of Purchase Orders", "Number of Receiving Records",
]


def build_multiple_batches(df: pd.DataFrame) -> pd.DataFrame:
    """Materials carrying more than one distinct Batch AND more than one
    distinct Batch SLED. This is an analysis view, not an issue flag — different
    production batches with different expiry dates may be perfectly normal."""
    rows: list[dict] = []
    has_sled = COL_BATCH_SLED in df.columns
    has_desc = COL_MATERIAL_DESC in df.columns
    has_supplier = COL_SUPPLIER_NAME in df.columns
    has_po = COL_PO in df.columns
    for material, g in df.groupby(COL_MATERIAL):
        n_batches = _nonblank(g[COL_BATCH]).nunique()
        dates = _nonblank(g[COL_BATCH_SLED]).dropna() if has_sled else pd.Series([], dtype="datetime64[ns]")
        n_dates = dates.nunique()
        if n_batches <= 1 or n_dates <= 1:
            continue
        rows.append({
            "Material": material,
            "Material Description": g[COL_MATERIAL_DESC].iloc[0] if has_desc else "",
            "Supplier Name": ", ".join(dict.fromkeys(
                v for v in g[COL_SUPPLIER_NAME].astype(str) if v.strip() and v.lower() != "nan"
            )) if has_supplier else "",
            "Number of Batch Values": int(n_batches),
            "Number of Expiry Dates": int(n_dates),
            "Earliest Expiry Date": dates.min() if not dates.empty else pd.NaT,
            "Latest Expiry Date": dates.max() if not dates.empty else pd.NaT,
            "Number of Purchase Orders": int(_nonblank(g[COL_PO]).nunique()) if has_po else 0,
            "Number of Receiving Records": int(len(g)),
        })
    out = pd.DataFrame(rows, columns=MULTI_BATCH_COLUMNS)
    if not out.empty:
        out = out.sort_values(
            ["Number of Expiry Dates", "Number of Batch Values"], ascending=False
        ).reset_index(drop=True)
    return out
