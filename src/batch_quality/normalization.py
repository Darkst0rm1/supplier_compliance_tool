"""Batch-string normalization and structural comparison helpers.

``normalize_batch`` builds a comparison-only key (uppercase, no spaces or
punctuation). The remaining helpers support the character-confusion rule and the
material batch/expiry pattern rule. None of these EVER mutate or replace the
original ``Batch`` value shown to the user — they only compute derived keys used
for detection.
"""
from __future__ import annotations

import re
from typing import Any, Optional

import pandas as pd


def normalize_batch(value: Any) -> str:
    """Comparison-only normalization: uppercase, drop spaces and separators
    (hyphens / slashes / periods / underscores and any other punctuation), keep
    letters and digits. ``31-5357`` / ``31 5357`` / ``31/5357`` -> ``315357``.
    Never replaces the original Batch value."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"[^0-9A-Z]", "", str(value).strip().upper())


# Characters commonly confused during label reading or manual entry.
CONFUSABLE_PAIRS = [
    ("I", "1"), ("O", "0"), ("S", "5"), ("B", "8"), ("Z", "2"), ("G", "6"),
]
_CONFUSABLE = {frozenset(p) for p in CONFUSABLE_PAIRS}

# How many confused-character positions two batches may differ by and still be
# treated as a probable mis-read of the same value.
MAX_CONFUSABLE_DIFFS = 2


def is_confusable(a: str, b: str) -> bool:
    """True if ``a`` and ``b`` are a commonly-confused character pair."""
    return a != b and frozenset((a, b)) in _CONFUSABLE


def confusable_diff_positions(a: str, b: str) -> Optional[list[int]]:
    """Positions where ``a`` and ``b`` differ — but only if they are the same
    length and EVERY differing position is a commonly-confused character pair.
    Otherwise ``None`` (not a probable mis-read)."""
    if len(a) != len(b) or a == b:
        return None
    diffs = [i for i, (x, y) in enumerate(zip(a, b)) if x != y]
    if not diffs:
        return None
    if all(is_confusable(a[i], b[i]) for i in diffs):
        return diffs
    return None


def batch_structure(s: str) -> str:
    """Coarse shape of a batch value: ``L`` letter, ``D`` digit, ``X`` other.
    ``OPOCT2227`` -> ``LLLLLDDDD``; ``2027AU13`` -> ``DDDDLLDD``."""
    out: list[str] = []
    for ch in s:
        if ch.isalpha():
            out.append("L")
        elif ch.isdigit():
            out.append("D")
        else:
            out.append("X")
    return "".join(out)


def structures_differ_significantly(a: str, b: str) -> bool:
    """True when two (normalized) batch values have meaningfully different
    shapes — different length, or a different letter/digit pattern. Used to tell
    a probable mis-key apart from two genuinely unrelated batch numbers."""
    if a == b:
        return False
    if len(a) != len(b):
        return True
    return batch_structure(a) != batch_structure(b)
