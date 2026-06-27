"""Batch Quality Analysis engine.

Reads a SAP receiving-history export and surfaces *possible* batch-number and
expiry-date data-quality issues for human review. Rule-based detection runs
first; the AI review agent then runs automatically over the issue groups as part
of file processing. The AI only explains and organizes a concern — it never
suggests corrected values or edits SAP data.

Self-contained subpackage (mirrors the standalone-engine pattern used by the
other reports in this app). Modules:

* ``loader``        — read the export, canonicalize headers, parse types.
* ``normalization`` — batch comparison key + structural/character helpers.
* ``rules``         — the seven rule-based detectors + the multi-batch summary.
* ``issue_groups``  — aggregate rule hits into reviewable issue groups + merge.
* ``ai_agent``      — automatic Anthropic-backed review agent (structured output).
* ``reviews``       — human-review record helpers (session-state backed).
* ``exporter``      — the three-sheet Excel workbook.
* ``page``          — the Streamlit page renderer.
"""
