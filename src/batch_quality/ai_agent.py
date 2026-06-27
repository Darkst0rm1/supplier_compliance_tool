"""Automatic AI review agent for batch-quality issue groups.

The agent runs as part of file processing — after the rule-based issue groups are
built — not as a separate per-issue button. For each issue group it reviews the
group summary plus only the relevant related/material/normalized-batch records
(never the whole workbook) and returns a validated structured review.

The AI is a **review assistant only**. It explains and organizes a concern to
help a human investigate. It must NOT suggest a corrected batch number or expiry
date, decide which record is correct, invent missing information, or modify SAP
data — and the structured output schema deliberately contains no corrected-value
fields.

Uses the official Anthropic Python SDK with structured outputs
(``client.messages.parse`` + a Pydantic model). Reads ``ANTHROPIC_API_KEY`` and
``ANTHROPIC_MODEL`` from the environment; if the key is missing the page keeps
every rule-based feature working and the AI columns read "Not Available".

Results are cached by a stable issue-group fingerprint so reruns don't re-send an
unchanged issue.
"""
from __future__ import annotations

import hashlib
import json
import os
from typing import Any, Callable, Literal, Optional

import pandas as pd
from pydantic import BaseModel

from .loader import NORMALIZED_BATCH

DEFAULT_MODEL = "claude-opus-4-8"

# Cap how many context records we send the AI — it reviews the issue group, not
# the whole workbook.
MAX_CONTEXT_RECORDS = 40

# Default cap on how many issue groups the agent reviews automatically (highest
# risk first). Protects against runaway cost on very large exports. Overridable
# via BATCH_QUALITY_MAX_AI_ISSUES.
DEFAULT_MAX_AI_ISSUES = 100


class AIReviewResult(BaseModel):
    """Structured AI review output. Intentionally has NO corrected/suggested/
    replacement value fields — the AI never proposes a fix."""

    ai_review_priority: Literal["High", "Medium", "Low"]
    review_summary: str
    pattern_identified: str
    reason_for_review: str
    recurring_pattern: str
    records_involved: str
    documents_to_verify: list[str]
    possible_root_causes: list[str]
    recommended_review_steps: list[str]
    questions_for_supplier_or_receiver: list[str]
    review_note: str


# Field names that must never appear on the AI output schema.
FORBIDDEN_FIELDS = {
    "suggested_batch", "suggested_expiry_date", "corrected_batch",
    "corrected_expiry_date", "replacement_value",
}

# Display/export column names, in order, mapped from the schema fields.
AI_COLUMNS = [
    "AI Review Priority",
    "AI Review Summary",
    "Pattern Identified",
    "Reason for Review",
    "Recurring Pattern",
    "Records Involved",
    "Documents to Verify",
    "Possible Root Causes",
    "Recommended Review Steps",
    "Questions for Supplier or Receiver",
    "AI Review Note",
]

SYSTEM_PROMPT = (
    "You are an AI review agent assisting an operations team with SAP receiving "
    "batch and shelf-life data.\n\n"
    "Review the supplied issue group and supporting records for traceability, "
    "recall readiness, inventory control and receiving-data quality.\n\n"
    "Do not suggest a corrected batch number.\n"
    "Do not suggest a corrected expiry date.\n"
    "Do not decide which record is correct.\n"
    "Do not invent missing information.\n"
    "Do not assume that different batches and different expiry dates are "
    "automatically wrong.\n\n"
    "Explain the pattern, why it deserves review, whether it appears recurring, "
    "what documents should be checked, possible root-cause categories, practical "
    "investigation steps, and questions for the supplier or receiver.\n\n"
    "Pay particular attention to character-reading issues such as I versus 1 and "
    "O versus 0, similar batch values, inconsistent batch formats, nearby "
    "receipts for the same material and supplier, and unusually different expiry "
    "dates.\n\n"
    "When the information is insufficient, clearly state that supplier labels, "
    "ASN information, packing slips, inbound-delivery data or receiving documents "
    "are required."
)


def get_api_key(override: Optional[str] = None) -> Optional[str]:
    return override or os.environ.get("ANTHROPIC_API_KEY")


def get_model(override: Optional[str] = None) -> str:
    return override or os.environ.get("ANTHROPIC_MODEL") or DEFAULT_MODEL


def is_ai_configured(api_key: Optional[str] = None) -> bool:
    return bool(get_api_key(api_key))


def get_max_issues() -> Optional[int]:
    raw = os.environ.get("BATCH_QUALITY_MAX_AI_ISSUES")
    if raw is None:
        return DEFAULT_MAX_AI_ISSUES
    raw = raw.strip().lower()
    if raw in ("", "0", "none", "all", "-1"):
        return None
    try:
        return int(raw)
    except ValueError:
        return DEFAULT_MAX_AI_ISSUES


def _records_to_list(df: pd.DataFrame, limit: Optional[int] = None) -> list[dict]:
    if df is None or df.empty:
        return []
    if limit is not None:
        df = df.head(limit)
    out: list[dict] = []
    for _, row in df.iterrows():
        rec: dict[str, Any] = {}
        for col, val in row.items():
            if pd.isna(val):
                rec[col] = None
            elif isinstance(val, pd.Timestamp):
                rec[col] = val.strftime("%Y-%m-%d")
            else:
                rec[col] = str(val)
        out.append(rec)
    return out


def build_ai_context(
    group_row: dict,
    related_records: pd.DataFrame,
    material_records: pd.DataFrame,
    normbatch_records: pd.DataFrame,
    max_records: int = MAX_CONTEXT_RECORDS,
) -> dict:
    """Assemble exactly the context the AI should see for one issue group:
    the group summary, its related records, and a capped set of relevant
    same-Material and same-Normalized-Batch records. Never the whole workbook."""
    return {
        "issue_group": {k: (None if v is None else str(v)) for k, v in group_row.items()},
        "related_records": _records_to_list(related_records),
        "material_records": _records_to_list(material_records, limit=max_records),
        "normalized_batch_records": _records_to_list(normbatch_records, limit=max_records),
    }


def issue_fingerprint(issue_type: str, related_records: pd.DataFrame) -> str:
    """Stable hash of an issue group from its issue type and the source values of
    its member rows. Lets cached AI results survive Streamlit reruns and only
    re-trigger when the underlying records change."""
    payload = _records_to_list(
        related_records.drop(columns=[NORMALIZED_BATCH], errors="ignore")
    )
    blob = json.dumps([issue_type, sorted(json.dumps(r, sort_keys=True, default=str) for r in payload)])
    return hashlib.sha1(blob.encode("utf-8")).hexdigest()


def review_issue(
    context: dict,
    api_key: Optional[str] = None,
    model: Optional[str] = None,
) -> AIReviewResult:
    """Call Claude to review one issue group and return validated structured
    output. ``context`` is the dict from :func:`build_ai_context`."""
    import anthropic  # lazy — keep the module importable without the dependency

    client = anthropic.Anthropic(api_key=get_api_key(api_key))
    user_content = (
        "Review this SAP batch / expiry-date data-quality concern. Base your "
        "response only on the records provided.\n\n"
        + json.dumps(context, indent=2, default=str)
    )
    response = client.messages.parse(
        model=get_model(model),
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}],
        output_format=AIReviewResult,
    )
    return response.parsed_output


def run_agent(
    result,
    related: pd.DataFrame,
    *,
    cache: Optional[dict] = None,
    api_key: Optional[str] = None,
    model: Optional[str] = None,
    max_issues: Optional[int] = None,
    progress_cb: Optional[Callable[[int, int], None]] = None,
) -> dict:
    """Automatically review the issue groups (highest risk first), caching by
    fingerprint. Returns a map Issue Group ID -> AIReviewResult for reviewed
    groups, or the sentinel ``"skipped"`` for groups beyond ``max_issues``.

    Groups already in ``cache`` (keyed by fingerprint) are reused without a new
    API call. ``related`` is unused directly but kept for signature symmetry.
    """
    cache = cache if cache is not None else {}
    df = result.df
    flagged = result.flagged
    ai_map: dict = {}
    if flagged.empty:
        return ai_map

    gids = list(flagged["Issue Group ID"])
    # Count groups that still need a real call (for the caller's progress bar).
    total = len(gids) if max_issues is None else min(len(gids), max_issues)
    done = 0
    attempted = 0
    ok = 0
    first_error: Optional[Exception] = None
    for n, gid in enumerate(gids):
        if max_issues is not None and n >= max_issues:
            ai_map[gid] = "skipped"
            continue
        idx = result.members.get(gid, [])
        related_records = df.loc[idx]
        row = flagged[flagged["Issue Group ID"] == gid].iloc[0]
        fp = issue_fingerprint(row["Issue Type"], related_records)
        if fp in cache:
            ai_map[gid] = cache[fp]
            done += 1
            ok += 1
            if progress_cb:
                progress_cb(done, total)
            continue

        mat = related_records["Material"].iloc[0] \
            if "Material" in df.columns and not related_records.empty else None
        material_records = df[df["Material"] == mat] if mat is not None else df.iloc[0:0]
        nbs = [v for v in related_records[NORMALIZED_BATCH].unique() if str(v).strip()] \
            if NORMALIZED_BATCH in related_records.columns else []
        normbatch_records = df[df[NORMALIZED_BATCH].isin(nbs)] if nbs else df.iloc[0:0]

        context = build_ai_context(row.to_dict(), related_records, material_records, normbatch_records)
        attempted += 1
        try:
            review = review_issue(context, api_key=api_key, model=model)
        except Exception as exc:  # noqa: BLE001 — tolerate a single bad group
            if first_error is None:
                first_error = exc
            ai_map[gid] = None
        else:
            cache[fp] = review
            ai_map[gid] = review
            ok += 1
        done += 1
        if progress_cb:
            progress_cb(done, total)

    # Systemic failure (e.g. bad credentials): surface it. Partial failures are
    # tolerated — those groups simply show no AI result.
    if attempted and ok == 0 and first_error is not None:
        raise first_error
    return ai_map


def ai_fields(review: Any, *, unavailable_label: str = "Not Available") -> dict:
    """Map an :class:`AIReviewResult` (or ``None`` / ``"skipped"``) to the
    AI_COLUMNS display dict."""
    if isinstance(review, AIReviewResult):
        return {
            "AI Review Priority": review.ai_review_priority,
            "AI Review Summary": review.review_summary,
            "Pattern Identified": review.pattern_identified,
            "Reason for Review": review.reason_for_review,
            "Recurring Pattern": review.recurring_pattern,
            "Records Involved": review.records_involved,
            "Documents to Verify": "; ".join(review.documents_to_verify),
            "Possible Root Causes": "; ".join(review.possible_root_causes),
            "Recommended Review Steps": "; ".join(review.recommended_review_steps),
            "Questions for Supplier or Receiver": "; ".join(review.questions_for_supplier_or_receiver),
            "AI Review Note": review.review_note,
        }
    label = "Not Reviewed (limit reached)" if review == "skipped" else unavailable_label
    out = {c: "" for c in AI_COLUMNS}
    out["AI Review Priority"] = label
    return out
