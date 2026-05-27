"""PO normalization helpers.

Rules (from spec):
 - Treat PO numbers as text
 - Trim spaces, remove line breaks, collapse internal whitespace
 - Remove trailing ".0" left behind by Excel float coercion
 - Keep leading zeros
 - Portal cells may carry multiple POs separated by , / ; newline or whitespace
 - Duplicates are removed for counting (caller's responsibility via .unique())
"""
from __future__ import annotations

import re
import pandas as pd

# Anything that could plausibly separate POs in a free-text portal cell.
_SPLIT_RE = re.compile(r"[,/;\n\r\t ]+")
_WHITESPACE_RE = re.compile(r"\s+")


def normalize_po(value) -> str:
    """Normalize a single PO value to a clean text token. Empty/NaN -> ''."""
    if value is None:
        return ""
    # Numeric coming out of pandas (e.g., 4500012345.0)
    if isinstance(value, float):
        if pd.isna(value):
            return ""
        if value.is_integer():
            value = int(value)
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    # Drop trailing ".0" from text like "4500012345.0"
    if s.endswith(".0"):
        try:
            f = float(s)
            if f.is_integer():
                s = str(int(f))
        except ValueError:
            pass
    # Remove all internal whitespace and line breaks
    s = _WHITESPACE_RE.sub("", s)
    return s


def split_multi_po(value) -> list[str]:
    """Split a portal cell that may contain multiple POs into normalized POs.

    Returns a de-duplicated list preserving first-seen order.
    """
    if value is None:
        return []
    if isinstance(value, float) and pd.isna(value):
        return []
    s = str(value).strip()
    if not s:
        return []
    parts = _SPLIT_RE.split(s)
    out: list[str] = []
    seen: set[str] = set()
    for p in parts:
        n = normalize_po(p)
        if n and n not in seen:
            seen.add(n)
            out.append(n)
    return out


def has_value(value) -> bool:
    """True if the cell holds a real non-empty value (used for Inbound Delivery)."""
    if value is None:
        return False
    if isinstance(value, float) and pd.isna(value):
        return False
    s = str(value).strip()
    return bool(s) and s.lower() != "nan"
