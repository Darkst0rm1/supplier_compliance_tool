"""AI-assisted review of a single batch-quality issue group.

The AI is a **review assistant only**. It explains and organizes the supplied
data-quality concern to help a human investigate. It must NOT suggest a
corrected batch number or expiry date, decide which record is correct, invent
missing information, or modify SAP data — and the structured output schema
deliberately contains no corrected-value fields.

Uses the official Anthropic Python SDK with structured outputs
(``client.messages.parse`` + a Pydantic model). Reads ``ANTHROPIC_API_KEY`` and
``ANTHROPIC_MODEL`` from the environment; if the key is missing the page keeps
every rule-based feature working and only disables the AI button.
"""
from __future__ import annotations

import json
import os
from typing import Any, Literal, Optional

import pandas as pd
from pydantic import BaseModel

from .loader import NORMALIZED_BATCH

DEFAULT_MODEL = "claude-opus-4-8"

# Cap how many context records we send the AI — it reviews the issue group, not
# the whole workbook.
MAX_CONTEXT_RECORDS = 40


class AIReviewResult(BaseModel):
    """Structured AI review output. Intentionally has NO corrected/suggested/
    replacement value fields — the AI never proposes a fix."""

    review_summary: str
    pattern_identified: str
    reason_for_review: str
    risk_level: Literal["High", "Medium", "Low"]
    records_involved: str
    recurring_pattern: str
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

SYSTEM_PROMPT = (
    "You are assisting with the review of SAP batch and expiry-date data for "
    "product traceability, inventory control and recall readiness.\n\n"
    "Your task is to explain the supplied data-quality concern and help a human "
    "reviewer investigate it.\n\n"
    "Do not suggest a corrected batch number. Do not suggest a corrected expiry "
    "date. Do not decide which record is correct. Do not invent missing "
    "information. Do not assume that multiple batches or expiry dates are "
    "automatically wrong.\n\n"
    "Compare the records, describe the pattern, explain why the issue requires "
    "review, identify repeated behavior, assign a review priority, list documents "
    "that should be checked, identify possible root-cause categories, and provide "
    "practical questions for the supplier or receiver.\n\n"
    "When the data is insufficient, clearly state that supplier labels, ASN "
    "information, packing slips, inbound deliveries or receiving documentation "
    "must be checked."
)


def get_api_key(override: Optional[str] = None) -> Optional[str]:
    return override or os.environ.get("ANTHROPIC_API_KEY")


def get_model(override: Optional[str] = None) -> str:
    return override or os.environ.get("ANTHROPIC_MODEL") or DEFAULT_MODEL


def is_ai_configured(api_key: Optional[str] = None) -> bool:
    return bool(get_api_key(api_key))


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
