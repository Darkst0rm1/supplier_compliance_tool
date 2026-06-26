"""Batch Quality Analysis engine.

Reads a SAP receiving-history export and surfaces *possible* batch-number and
expiry-date data-quality issues for human review. Rule-based detection runs
first; an AI assistant (optional) only explains and organizes a single selected
issue group — it never suggests corrected values or edits SAP data.

Self-contained subpackage (mirrors the standalone-engine pattern used by the
other reports in this app). Modules:

* ``loader``       — read + normalize the export.
* ``rules``        — the six rule-based detectors + the multi-batch summary.
* ``issue_groups`` — aggregate rule hits into reviewable issue groups.
* ``reviews``      — human-review record helpers (session-state backed).
* ``ai_review``    — Anthropic-backed review assistant (structured output).
* ``exporter``     — the four-sheet Excel workbook.
"""
