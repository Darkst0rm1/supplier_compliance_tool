"""Rule-based batch-quality detectors.

Each rule scans the normalized receiving DataFrame and returns ``RuleHit``s — an
issue type/risk plus the original SAP row indices involved. No row is ever
modified or deleted; the rules only *identify* records for human review. AI is
not involved here.

Rules (from the business spec):

1. Probable same batch entered differently        -> Possible Batch Format Variation (Medium)
2. Probable same batch, conflicting expiry dates   -> Possible Matching Batch With Conflicting Expiry Dates (High)
3. Probable character-entry error (I/1, O/0, …)     -> Possible Character Entry Error (Medium)
4. Same material + supplier received close together
   with unusual batch & expiry differences         -> Material Batch and Expiry Pattern Conflict (High/Medium)
5. Duplicate PO+Material+Batch+SLED combination     -> Duplicate Receiving Combination (Medium)
6. Completely identical row                          -> Exact Duplicate Record (High)
7. Batch formatting concern (spaces/punct/short/…)  -> Batch Format Review (Low)

Plus an analysis-only summary of materials carrying more than one batch AND more
than one expiry date (NOT an issue flag).
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
    COL_SUPPLIER,
    COL_SUPPLIER_NAME,
    COL_VENDOR,
    NORMALIZED_BATCH,
)
from .normalization import (
    MAX_CONFUSABLE_DIFFS,
    confusable_diff_positions,
    structures_differ_significantly,
)

# ---------------------------------------------------------------------------
# Issue types and risk levels
# ---------------------------------------------------------------------------
ISSUE_FORMAT_VARIATION = "Possible Batch Format Variation"
ISSUE_CONFLICTING_EXPIRY = "Possible Matching Batch With Conflicting Expiry Dates"
ISSUE_CHARACTER_ENTRY = "Possible Character Entry Error"
ISSUE_PATTERN_CONFLICT = "Material Batch and Expiry Pattern Conflict"
ISSUE_DUPLICATE_RECEIVING = "Duplicate Receiving Combination"
ISSUE_EXACT_DUPLICATE = "Exact Duplicate Record"
ISSUE_FORMAT_REVIEW = "Batch Format Review"

RISK_HIGH = "High"
RISK_MEDIUM = "Medium"
RISK_LOW = "Low"
RISK_ORDER = {RISK_HIGH: 0, RISK_MEDIUM: 1, RISK_LOW: 2}

# Rule 4 tuning.
CLOSE_PERIOD_DAYS = 14            # "received close together" window
EXPIRY_SIGNIFICANT_DAYS = 180    # min expiry gap to call it significant
EXPIRY_HIGH_RISK_DAYS = 365      # gap above which the pair is High risk

# Performance guard for the character-confusion rule.
_MAX_BATCHES_PER_LENGTH = 400


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
    """Group by Material + Normalized Batch. More than one distinct nonblank
    Batch SLED -> conflicting expiry (High, Rule 2). Else more than one original
    Batch format -> format variation (Medium, Rule 1)."""
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
                f"Material {material}: a probable matching batch ({nb}) is represented "
                f"with more than one expiry date {dates} across formats {formats} and "
                "requires document verification.",
                list(g.index),
            ))
        elif len(formats) > 1:
            hits.append(RuleHit(
                ISSUE_FORMAT_VARIATION, RISK_MEDIUM,
                f"Material {material}: probable same batch {nb} entered in "
                f"{len(formats)} different formats {formats}.",
                list(g.index),
            ))
    return hits


# ---------------------------------------------------------------------------
# Rule 3: probable character-entry error (I/1, O/0, S/5, B/8, Z/2, G/6)
# ---------------------------------------------------------------------------
def rule_character_entry(df: pd.DataFrame) -> list[RuleHit]:
    """Within a Material, two batch values of the same length that differ only
    by commonly-confused characters (e.g. ``NI8D61`` vs ``N18D61``) are flagged
    for review. Comparison uses the normalized batch so punctuation/casing don't
    create false differences. Never declares which value is correct."""
    hits: list[RuleHit] = []
    sub = df[df[NORMALIZED_BATCH].astype(str).str.strip() != ""]
    for material, g in sub.groupby(COL_MATERIAL):
        norm_to_idx: dict[str, list] = {}
        norm_to_orig: dict[str, str] = {}
        for idx, nb, orig in zip(g.index, g[NORMALIZED_BATCH], g[COL_BATCH]):
            nb = str(nb)
            norm_to_idx.setdefault(nb, []).append(idx)
            norm_to_orig.setdefault(nb, str(orig).strip())
        norms = list(norm_to_idx)
        if len(norms) < 2:
            continue
        by_len: dict[int, list[str]] = {}
        for nb in norms:
            by_len.setdefault(len(nb), []).append(nb)
        seen: set = set()
        for length, bucket in by_len.items():
            if length == 0 or len(bucket) > _MAX_BATCHES_PER_LENGTH:
                continue
            for i in range(len(bucket)):
                for j in range(i + 1, len(bucket)):
                    a, b = bucket[i], bucket[j]
                    diffs = confusable_diff_positions(a, b)
                    if diffs is None or len(diffs) > MAX_CONFUSABLE_DIFFS:
                        continue
                    key = frozenset((a, b))
                    if key in seen:
                        continue
                    seen.add(key)
                    idxs = sorted(set(norm_to_idx[a] + norm_to_idx[b]))
                    hits.append(RuleHit(
                        ISSUE_CHARACTER_ENTRY, RISK_MEDIUM,
                        f"Material {material}: batch values "
                        f"{norm_to_orig[a]!r} and {norm_to_orig[b]!r} are highly similar "
                        "and differ by a character commonly confused during label "
                        "reading or manual entry.",
                        idxs,
                    ))
    return hits


# ---------------------------------------------------------------------------
# Rule 4: same material + supplier received close together with unusual
# batch-structure and expiry differences
# ---------------------------------------------------------------------------
def rule_pattern_conflict(
    df: pd.DataFrame,
    window_days: int = CLOSE_PERIOD_DAYS,
    expiry_days: int = EXPIRY_SIGNIFICANT_DAYS,
) -> list[RuleHit]:
    """Compare records sharing Material + Supplier (or Vendor) received within a
    close period. Flag a pair whose batch structures differ significantly AND
    whose expiry dates differ significantly — a possible inconsistent receipt or
    mixed-up lot. Does not decide which record is correct."""
    if COL_BATCH_SLED not in df.columns or COL_RECEIVED not in df.columns:
        return []
    skey = COL_SUPPLIER if COL_SUPPLIER in df.columns else (
        COL_VENDOR if COL_VENDOR in df.columns else None
    )
    if skey is None:
        return []

    mask = (
        df[COL_RECEIVED].notna()
        & df[COL_BATCH_SLED].notna()
        & (df[COL_BATCH].astype(str).str.strip() != "")
        & (df[skey].astype(str).str.strip() != "")
    )
    sub = df[mask]
    hits: list[RuleHit] = []
    for (material, supplier), grp in sub.groupby([COL_MATERIAL, skey]):
        if len(grp) < 2:
            continue
        g = grp.sort_values(COL_RECEIVED)
        idx = list(g.index)
        norms = [str(v) for v in g[NORMALIZED_BATCH]]
        origs = [str(v).strip() for v in g[COL_BATCH]]
        sleds = list(g[COL_BATCH_SLED])
        recvs = list(g[COL_RECEIVED])
        n = len(idx)
        seen: set = set()
        for i in range(n):
            for j in range(i + 1, n):
                dr = abs((recvs[j] - recvs[i]).days)
                if dr > window_days:
                    break  # sorted by received date — no closer pair beyond j
                if norms[i] == norms[j]:
                    continue
                if not structures_differ_significantly(norms[i], norms[j]):
                    continue
                de = abs((sleds[j] - sleds[i]).days)
                if de < expiry_days:
                    continue
                pair = tuple(sorted((idx[i], idx[j])))
                if pair in seen:
                    continue
                seen.add(pair)
                risk = RISK_HIGH if de > EXPIRY_HIGH_RISK_DAYS else RISK_MEDIUM
                hits.append(RuleHit(
                    ISSUE_PATTERN_CONFLICT, risk,
                    f"Material {material}: the same material and supplier were received "
                    f"within {dr} day(s), but batch values {origs[i]!r} "
                    f"(expiry {pd.Timestamp(sleds[i]):%Y-%m-%d}) and {origs[j]!r} "
                    f"(expiry {pd.Timestamp(sleds[j]):%Y-%m-%d}) have significantly "
                    "different structures and expiry dates. Review whether these "
                    "represent valid separate production lots or inconsistent "
                    "receiving data.",
                    list(pair),
                ))
    return hits


# ---------------------------------------------------------------------------
# Rule 5: duplicate receiving combination
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
# Rule 6: exact duplicate row
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
# Rule 7: batch formatting concern (per distinct Material+Batch)
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
    """Run every rule (1–7) and return all hits."""
    return (
        rule_batch_groups(df)
        + rule_character_entry(df)
        + rule_pattern_conflict(df)
        + rule_duplicate_receiving(df)
        + rule_exact_duplicate(df)
        + rule_format_review(df)
    )


# ---------------------------------------------------------------------------
# Analysis only: materials with multiple batches AND multiple dates
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
