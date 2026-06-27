"""Batch Quality Analysis — surface possible SAP batch-number and expiry-date
data-quality issues for human review.

Upload the SAP receiving-history export, click *Process Batch Quality File*, and
the page runs rule-based detection, builds issue groups, and runs the AI review
agent automatically over those groups. The AI never suggests corrected values or
edits SAP data. The page logic lives in ``src/batch_quality/page.py``.
"""
from __future__ import annotations

from src.batch_quality.page import render

render()
