"""Human review records for batch-quality issue groups.

First version stores reviews in Streamlit ``session_state`` (no database — the
rest of the app's persistence is the optional Neon DB for column variants only).
This module is pure dict logic so it's testable without Streamlit; the page
passes ``st.session_state`` (or a plain dict) as the ``store``.
"""
from __future__ import annotations

from typing import Any

CONFIRMED_OPTIONS = ["Yes", "No", "Needs Investigation"]
ROOT_CAUSE_OPTIONS = [
    "Supplier-provided batch information",
    "Receiving data entry",
    "Batch label misread",
    "Expiry-date entry",
    "Batch formatting inconsistency",
    "Internal transfer",
    "Duplicate receipt",
    "Valid separate batches",
    "Unable to determine",
]
RESPONSIBLE_AREA_OPTIONS = [
    "Supplier",
    "Tree of Life Receiving",
    "Lineage Receiving",
    "Master Data",
    "SAP Process",
    "Unknown",
]
REVIEW_STATUS_OPTIONS = ["Not Reviewed", "In Review", "Waiting for Documents", "Completed"]
FOLLOW_UP_OPTIONS = ["Yes", "No"]

# The human-review fields stored per issue group.
HUMAN_REVIEW_FIELDS = [
    "Confirmed Issue",
    "Root Cause",
    "Responsible Area",
    "Review Status",
    "Follow-Up Required",
    "Reviewer Comment",
]


def save_review(store: dict, group_id: str, fields: dict[str, Any]) -> dict:
    """Record (or overwrite) the human review for ``group_id`` in ``store``.

    ``store`` maps Issue Group ID -> review dict. Returns the same store for
    convenience. Only known fields are kept."""
    review = {k: fields.get(k, "") for k in HUMAN_REVIEW_FIELDS}
    store[group_id] = review
    return store


def get_review(store: dict, group_id: str) -> dict:
    """The saved review for ``group_id``, or an empty template if none saved."""
    return store.get(group_id, {k: "" for k in HUMAN_REVIEW_FIELDS})
