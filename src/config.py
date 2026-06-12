"""Centralized column names and status codes used across the importers and engine."""

# Canonical SAP columns the engine expects after import. Real SAP exports use
# different headers (Plant, Vendor, Appt. Date) — the importer renames them
# before any other processing.
SAP_CANONICAL_COLUMNS = [
    "PO Number",
    "Vendor Number",
    "Vendor Name",
    "Warehouse",
    "PO Status",
    "Appointment Date",
    "Delivery Date",
    "Inbound Delivery",
    "Inbound Delivery Status",
]

# Only these two are *strictly* required for the tool to function. Everything
# else is filled with blanks if absent, so older or partial SAP exports still
# load (with empty warehouse/supplier rollups noted in the summary).
SAP_HARD_REQUIRED_COLUMNS = [
    "PO Number",
    "Inbound Delivery",
]

# Real SAP export header -> canonical header used by the engine.
SAP_COLUMN_ALIASES = {
    "Plant": "Warehouse",
    "Vendor": "Vendor Number",
    "Appt. Date": "Appointment Date",
}

# Columns the SAP export *may* include. If absent the tool fills them blank
# rather than erroring. Currently used only for the SAP Export Data sheet.
SAP_OPTIONAL_DATE_COLUMNS = [
    "Confirmed PU Date",
    "Est PU Date",
]

PORTAL_REQUIRED_COLUMNS = [
    "PO Number",
    "Supplier Name",
    "Upload Date",
]

# Optional portal columns. If present they flow through into the report; if
# absent the engine treats every portal upload as compliant (back-compat).
PORTAL_OPTIONAL_COLUMNS = [
    "File Name",
    "Uploaded By",
    "File Status",
    "Downloaded By",
    "Download Date",
    "Invalid Comment",
]

# Aliases for portal headers (some exports may use "PO Number(s)" / "Supplier").
PORTAL_COLUMN_ALIASES = {
    "PO Number(s)": "PO Number",
    "Supplier": "Supplier Name",
}

# File Status semantics:
#   Approved   -> file was reviewed and accepted by TOL
#   Received   -> file received/accepted by TOL (counts as compliant, like Approved)
#   Submitted  -> uploaded, awaiting TOL review (still counts as compliant
#                 because the supplier did their part on time)
#   Invalid    -> uploaded but rejected by TOL (does NOT count as compliant)
#   (blank)    -> back-compat: treat as Approved
PORTAL_STATUS_APPROVED = "Approved"
PORTAL_STATUS_RECEIVED = "Received"
PORTAL_STATUS_SUBMITTED = "Submitted"
PORTAL_STATUS_INVALID = "Invalid"
PORTAL_VALID_STATUSES = {
    PORTAL_STATUS_APPROVED,
    PORTAL_STATUS_RECEIVED,
    PORTAL_STATUS_SUBMITTED,
    "",
}
PORTAL_PENDING_STATUSES = {PORTAL_STATUS_SUBMITTED}

# Bill-back: suppliers are charged a flat fee for every inbound PO whose
# documentation was never uploaded to the portal (a "Missing Inbound Document").
# Uploaded-but-Invalid POs are NOT billed -- the supplier attempted.
BILLBACK_FEE_PER_OCCURRENCE = 200          # USD per missing inbound document
BILLBACK_REASON = "Missing Inbound Document"

# SAP Inbound Delivery Status codes (real export uses C / A / B; spec said
# "P" for processing but actual SAP exports use "B"). Both are recognized.
PO_STATUS_CLOSED = "C"
PO_STATUS_APPROVED = "A"
PO_STATUS_PROCESSING_CODES = {"B", "P"}

STATUS_LABELS = {
    "C": "Closed / Review Separately",
    "A": "Approved",
    "B": "Processing / In-Progress",
    "P": "Processing / Pending",
}

# SAP date columns that scope the report month. A PO is "in scope" if ANY of
# these dates falls in the selected month (union). Hardcoded as a business
# rule for auditability -- no UI choice. Order is purely for the audit string
# printed in the Monthly Summary.
SAP_FILTER_DATE_COLUMNS = [
    "Delivery Date",
    "Appointment Date",
    "Confirmed PU Date",
    "Est PU Date",
]

# Compliance labels used in the report
COMPLIANT = "Submitted / Compliant"
NON_COMPLIANT = "Missing Portal File / Non-Compliant"
PORTAL_NOT_IN_SAP = "Portal PO Not Found In SAP / Needs Review"
PORTAL_NO_SAP_INBOUND = "Portal File Exists But No SAP Inbound / Needs Review"

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
