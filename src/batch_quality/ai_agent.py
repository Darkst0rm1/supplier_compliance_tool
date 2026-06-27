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
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from typing import Any, Callable, Literal, Optional

import pandas as pd
from pydantic import BaseModel

from .loader import NORMALIZED_BATCH

# Haiku is the default: this dashboard makes many small "explain why this was
# flagged" calls, where a fast, inexpensive model is the right fit. Override with
# ANTHROPIC_MODEL (e.g. claude-opus-4-8) for deeper reviews.
DEFAULT_MODEL = "claude-haiku-4-5"

# Cap how many context records we send the AI — it reviews the issue group, not
# the whole workbook.
MAX_CONTEXT_RECORDS = 40

# Default cap on how many issue groups the agent reviews automatically (highest
# risk first). Protects against runaway cost/time on very large exports.
# Overridable via BATCH_QUALITY_MAX_AI_ISSUES.
DEFAULT_MAX_AI_ISSUES = 25

# How many AI calls run concurrently. Overridable via BATCH_QUALITY_AI_CONCURRENCY.
DEFAULT_AI_CONCURRENCY = 8

# --- Pluggable backend ------------------------------------------------------
# Default backend is Anthropic. Set BATCH_QUALITY_LLM_BASE_URL to use any
# OpenAI-compatible endpoint instead (free options: Hugging Face router, Groq,
# Ollama local, OpenRouter). Then also set BATCH_QUALITY_LLM_MODEL and, unless
# the endpoint needs no auth (e.g. local Ollama), BATCH_QUALITY_LLM_API_KEY
# (HF_TOKEN is accepted as a fallback for the Hugging Face router).
#
#   Hugging Face:  BASE_URL=https://router.huggingface.co/v1   key=hf_...   model=meta-llama/Llama-3.1-8B-Instruct
#   Groq:          BASE_URL=https://api.groq.com/openai/v1      key=gsk_...  model=llama-3.3-70b-versatile
#   Ollama (local):BASE_URL=http://localhost:11434/v1           key=(none)   model=llama3.1
#   OpenRouter:    BASE_URL=https://openrouter.ai/api/v1        key=sk-or-.. model=...:free
DEFAULT_OPENAI_MODEL = "meta-llama/Llama-3.1-8B-Instruct"
OPENAI_REQUEST_TIMEOUT = 120


def get_backend() -> str:
    """``"openai"`` when an OpenAI-compatible base URL is configured, else
    ``"anthropic"`` (the default)."""
    return "openai" if os.environ.get("BATCH_QUALITY_LLM_BASE_URL") else "anthropic"


def _openai_config() -> tuple[str, str]:
    base = os.environ.get("BATCH_QUALITY_LLM_BASE_URL", "").rstrip("/")
    key = os.environ.get("BATCH_QUALITY_LLM_API_KEY") or os.environ.get("HF_TOKEN") or ""
    return base, key


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
    if override:
        return override
    if get_backend() == "openai":
        return os.environ.get("BATCH_QUALITY_LLM_MODEL") or DEFAULT_OPENAI_MODEL
    return os.environ.get("ANTHROPIC_MODEL") or DEFAULT_MODEL


def is_ai_configured(api_key: Optional[str] = None) -> bool:
    if get_backend() == "openai":
        # A base URL is enough (local Ollama needs no key); a key is required
        # only for hosted providers, which is enforced by the call itself.
        return bool(os.environ.get("BATCH_QUALITY_LLM_BASE_URL"))
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


def get_concurrency() -> int:
    raw = os.environ.get("BATCH_QUALITY_AI_CONCURRENCY")
    if raw is None:
        return DEFAULT_AI_CONCURRENCY
    try:
        return max(1, int(raw.strip()))
    except ValueError:
        return DEFAULT_AI_CONCURRENCY


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


_USER_PREFIX = (
    "Review this SAP batch / expiry-date data-quality concern. Base your "
    "response only on the records provided.\n\n"
)


def review_issue(
    context: dict,
    api_key: Optional[str] = None,
    model: Optional[str] = None,
) -> AIReviewResult:
    """Review one issue group and return validated structured output. Dispatches
    to the configured backend (Anthropic by default, otherwise any
    OpenAI-compatible endpoint). ``context`` is from :func:`build_ai_context`."""
    if get_backend() == "openai":
        return _review_via_openai(context, model=model)

    import anthropic  # lazy — keep the module importable without the dependency

    client = anthropic.Anthropic(api_key=get_api_key(api_key))
    response = client.messages.parse(
        model=get_model(model),
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": _USER_PREFIX + json.dumps(context, indent=2, default=str)}],
        output_format=AIReviewResult,
    )
    return response.parsed_output


def _extract_json(text: str) -> str:
    """Pull the first JSON object out of a model response (tolerates ```json
    fences and surrounding prose from less strict models)."""
    t = text.strip()
    if t.startswith("```"):
        t = t[3:]
        if t[:4].lower() == "json":
            t = t[4:]
        t = t.split("```", 1)[0].strip()
    start = t.find("{")
    if start == -1:
        return t
    depth = 0
    for i in range(start, len(t)):
        if t[i] == "{":
            depth += 1
        elif t[i] == "}":
            depth -= 1
            if depth == 0:
                return t[start:i + 1]
    return t[start:]


def _field_guide() -> str:
    """Concise per-field instruction derived from the schema. Lighter-weight than
    dumping the raw JSON schema, which weaker models tend to echo back verbatim."""
    import typing
    lines = []
    for name, field in AIReviewResult.model_fields.items():
        ann = field.annotation
        origin = typing.get_origin(ann)
        if origin is list:
            kind = "array of short strings"
        elif origin is typing.Literal:
            kind = "one of " + ", ".join(repr(a) for a in typing.get_args(ann))
        else:
            kind = "string"
        lines.append(f'  "{name}": {kind}')
    return "\n".join(lines)


def _example_object() -> str:
    """A filled-in example so the model returns VALUES, not the schema."""
    return json.dumps({
        "ai_review_priority": "Medium",
        "review_summary": "<one short paragraph>",
        "pattern_identified": "<the pattern across the records>",
        "reason_for_review": "<why it needs review>",
        "recurring_pattern": "<does it repeat? brief>",
        "records_involved": "<count / which records>",
        "documents_to_verify": ["<doc 1>", "<doc 2>"],
        "possible_root_causes": ["<cause 1>", "<cause 2>"],
        "recommended_review_steps": ["<step 1>", "<step 2>"],
        "questions_for_supplier_or_receiver": ["<question 1>"],
        "review_note": "<concise note>",
    }, indent=2)


def _review_via_openai(context: dict, model: Optional[str] = None) -> AIReviewResult:
    """Review one issue group via an OpenAI-compatible chat-completions endpoint
    (Hugging Face router, Groq, Ollama, OpenRouter, …). Requests JSON output and
    validates it against :class:`AIReviewResult`, with one corrective retry."""
    import requests  # lazy

    base_url, key = _openai_config()
    if not base_url:
        raise RuntimeError("BATCH_QUALITY_LLM_BASE_URL is not set.")
    system = (
        SYSTEM_PROMPT
        + "\n\nRespond with ONLY a single JSON object containing EXACTLY these keys "
        "with these value types (no prose, no markdown, do NOT return the schema "
        "itself, fill in real values):\n" + _field_guide()
        + "\n\nExample of the expected shape (replace the placeholder values):\n"
        + _example_object()
    )
    user = _USER_PREFIX + json.dumps(context, indent=2, default=str)
    messages = [{"role": "system", "content": system},
                {"role": "user", "content": user}]
    headers = {"Content-Type": "application/json"}
    if key:
        headers["Authorization"] = f"Bearer {key}"
    url = f"{base_url}/chat/completions"

    def _call(msgs, use_json_format=True):
        payload = {"model": model or get_model(), "messages": msgs,
                   "max_tokens": 4096, "temperature": 0}
        if use_json_format:
            payload["response_format"] = {"type": "json_object"}
        resp = requests.post(url, json=payload, headers=headers, timeout=OPENAI_REQUEST_TIMEOUT)
        if resp.status_code >= 400 and use_json_format:
            # Some models/providers reject response_format — retry without it.
            payload.pop("response_format", None)
            resp = requests.post(url, json=payload, headers=headers, timeout=OPENAI_REQUEST_TIMEOUT)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]

    content = _call(messages)
    try:
        return AIReviewResult.model_validate_json(_extract_json(content))
    except Exception:  # noqa: BLE001 — one corrective retry for weaker models
        messages.append({"role": "assistant", "content": content})
        messages.append({"role": "user", "content": (
            "That was not valid. Return ONLY the JSON object with the exact keys "
            "listed, filled with real values for this issue — not the schema."
        )})
        content = _call(messages)
        return AIReviewResult.model_validate_json(_extract_json(content))


def run_agent(
    result,
    related: pd.DataFrame,
    *,
    cache: Optional[dict] = None,
    api_key: Optional[str] = None,
    model: Optional[str] = None,
    max_issues: Optional[int] = None,
    concurrency: Optional[int] = None,
    progress_cb: Optional[Callable[[int, int], None]] = None,
) -> dict:
    """Automatically review the issue groups (highest risk first), caching by
    fingerprint. Returns a map Issue Group ID -> AIReviewResult for reviewed
    groups, or the sentinel ``"skipped"`` for groups beyond ``max_issues``.

    Uncached groups are reviewed concurrently (``concurrency`` worker threads;
    the Anthropic SDK is HTTP/IO-bound, so threads parallelize the API latency).
    Groups already in ``cache`` (keyed by fingerprint) are reused without a new
    API call. ``related`` is unused directly but kept for signature symmetry.
    """
    cache = cache if cache is not None else {}
    workers = concurrency if concurrency is not None else get_concurrency()
    df = result.df
    flagged = result.flagged
    ai_map: dict = {}
    if flagged.empty:
        return ai_map

    gids = list(flagged["Issue Group ID"])
    in_scope = gids if max_issues is None else gids[:max_issues]
    for gid in gids[len(in_scope):]:
        ai_map[gid] = "skipped"

    total = len(in_scope)
    cached_ok = 0
    pending: list[tuple[str, str, dict]] = []  # (gid, fingerprint, context)
    for gid in in_scope:
        idx = result.members.get(gid, [])
        related_records = df.loc[idx]
        row = flagged[flagged["Issue Group ID"] == gid].iloc[0]
        fp = issue_fingerprint(row["Issue Type"], related_records)
        if fp in cache:
            ai_map[gid] = cache[fp]
            cached_ok += 1
            continue
        mat = related_records["Material"].iloc[0] \
            if "Material" in df.columns and not related_records.empty else None
        material_records = df[df["Material"] == mat] if mat is not None else df.iloc[0:0]
        nbs = [v for v in related_records[NORMALIZED_BATCH].unique() if str(v).strip()] \
            if NORMALIZED_BATCH in related_records.columns else []
        normbatch_records = df[df[NORMALIZED_BATCH].isin(nbs)] if nbs else df.iloc[0:0]
        context = build_ai_context(row.to_dict(), related_records, material_records, normbatch_records)
        pending.append((gid, fp, context))

    done = cached_ok
    if progress_cb and total:
        progress_cb(done, total)

    new_ok = 0
    first_error: Optional[Exception] = None
    if pending:
        lock = Lock()
        with ThreadPoolExecutor(max_workers=min(workers, len(pending))) as pool:
            futures = {
                pool.submit(review_issue, ctx, api_key, model): (gid, fp)
                for gid, fp, ctx in pending
            }
            for fut in as_completed(futures):
                gid, fp = futures[fut]
                try:
                    review = fut.result()
                except Exception as exc:  # noqa: BLE001 — tolerate a single bad group
                    with lock:
                        if first_error is None:
                            first_error = exc
                        ai_map[gid] = None
                else:
                    with lock:
                        cache[fp] = review
                        ai_map[gid] = review
                        new_ok += 1
                with lock:
                    done += 1
                    if progress_cb:
                        progress_cb(done, total)

    # Systemic failure (e.g. bad credentials): every actual API call failed and
    # nothing was served from cache. Surface it. Partial failures are tolerated.
    if pending and new_ok == 0 and cached_ok == 0 and first_error is not None:
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
