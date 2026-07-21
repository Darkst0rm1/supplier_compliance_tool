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

# PO numbers beginning with any of these prefixes are disregarded entirely —
# they're a different PO type that is not subject to portal inbound
# documentation, so they must never count toward compliance. Excluded from
# BOTH the SAP and Portal sides during import.
EXCLUDED_PO_PREFIXES = ("6",)

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

# ---------------------------------------------------------------------------
# Receiving Log (optional third input)
# ---------------------------------------------------------------------------
# The dock receiving log records what the receiver physically observed. It
# answers a question the portal structurally cannot: not "was a file uploaded"
# but "did the document's contents match the goods on the truck".
#
# Document accuracy is reported SEPARATELY from compliance. It never changes
# the Compliance Percentage or the bill-back -- those stay defined purely by
# portal file presence, so the audited number keeps its existing meaning.
RECEIVING_CANONICAL_COLUMNS = [
    "Receiving Date",
    "PO Number",
    "Carrier",
    "Inbound File Received",
    "Correct Batch",
    "Correct BBD",
    "Correct QTY",
    "Results of Inspection",
    "Receiver Initials",
    "Comments",
]

# Real receiving-log headers drift in spacing ("Y / N" vs "Y/N") and the sheet
# schema changed mid-year. Aliases are matched on a whitespace-stripped,
# lowercased form of the header so spacing never breaks the import.
RECEIVING_COLUMN_ALIASES = {
    "date": "Receiving Date",
    "po#": "PO Number",
    "po": "PO Number",
    "ponumber": "PO Number",
    "ponumber(s)": "PO Number",
    "carrier": "Carrier",
    "inboundfiley/n": "Inbound File Received",
    "inboundfile": "Inbound File Received",
    "correctbatchreceivedy/n": "Correct Batch",
    "correctbatchreceived": "Correct Batch",
    "correctbbdreceivedy/n": "Correct BBD",
    "correctbbdreceived": "Correct BBD",
    "correctqtyreceivedy/n": "Correct QTY",
    "correctqtyreceived": "Correct QTY",
    "resultsofinspection": "Results of Inspection",
    "receiverinitials": "Receiver Initials",
    "comments": "Comments",
}

# A receiving-log sheet is only usable if it carries at least one of these.
# The Jan-Apr sheets use the older 14-column schema that predates the audit
# columns entirely -- those sheets are skipped, not treated as all-blank.
RECEIVING_AUDIT_COLUMNS = [
    "Inbound File Received",
    "Correct Batch",
    "Correct BBD",
    "Correct QTY",
]

# The three columns that make up the document-accuracy check.
RECEIVING_ACCURACY_COLUMNS = ["Correct Batch", "Correct BBD", "Correct QTY"]

# The dock log's PO cell is hand-typed and mixes in references that are not
# SAP POs at all: carrier refs (TR-34306), supplier PO numbers (GHPO-23467),
# free text ("Return", "SILANI"), and short internal numbers (7176). SAP POs
# are purely numeric and at least this long. Anything else is dropped from the
# join and counted, so the log's real coverage is visible rather than implied.
RECEIVING_MIN_PO_DIGITS = 7

# Free-text Y/N answers normalize to exactly these, or "" for unanswered.
RECEIVING_YES = "YES"
RECEIVING_NO = "NO"
RECEIVING_YES_VALUES = {"YES", "Y", "YES ", "TRUE", "1"}
RECEIVING_NO_VALUES = {"NO", "N", "FALSE", "0"}

# Compliance labels used in the report
COMPLIANT = "Submitted / Compliant"
NON_COMPLIANT = "Missing Portal File / Non-Compliant"
PORTAL_NOT_IN_SAP = "Portal PO Not Found In SAP / Needs Review"
PORTAL_NO_SAP_INBOUND = "Portal File Exists But No SAP Inbound / Needs Review"

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

# --- Supplier exceptions ---------------------------------------------------
# Suppliers with an approved exception are not required to upload inbound
# documentation. The list is sourced from the Master Inbound Delivery Compliance
# Tracker workbook, as the union of two lists:
#   1. Tracker sheet, Compliance Status == TRACKER_STATUS_UNABLE_TO_COMPLY (24)
#   2. "POs received" sheet, a column literally containing "EXEMPT" (3)
# The two overlap on 2 suppliers, so the union is 25.
TRACKER_SHEET = "Tracker"
TRACKER_STATUS_COLUMN = "Compliance Status"
TRACKER_NAME_COLUMN = "Supplier Names "  # NB: trailing space, as in the workbook

# NB: DOUBLE space after the dash. This is the literal value in the workbook --
# do not "correct" it. The Summary sheet words it differently ("NO - Unable to
# comply - Approved exceptions"); that sheet is not the source of truth.
TRACKER_STATUS_UNABLE_TO_COMPLY = "NO -  Unable to Comply"

TRACKER_EXEMPT_SHEET = "POs received"
TRACKER_EXEMPT_MARKER = "EXEMPT"

REASON_UNABLE_TO_COMPLY = "Unable to Comply"
REASON_EXEMPT_MARK = "EXEMPT mark"
REASON_MANUAL = "Manual"

# Exception Status values shown on the Supplier Summary sheet.
EXCEPTION_STATUS_EXCEPTION = "Exception"
EXCEPTION_STATUS_EXPECTED = "Expected to upload"
EXCEPTION_STATUS_NOT_ON_TRACKER = "Not on tracker"
