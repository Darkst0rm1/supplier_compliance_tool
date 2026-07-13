"""Apply the compliance rules and build every report dataframe.

The engine treats one (normalized) PO as the unit of analysis. SAP rows are
de-duplicated by Normalized PO Number for the per-PO sheets so that line-item
duplicates don't inflate the report. Counts use .nunique() throughout.

SAP rows are first scoped to the report month using a user-chosen date column
(Delivery Date by default) so a multi-year SAP export doesn't dilute the
compliance percentage.

Portal File Status rules:
 - Approved / Submitted / (blank)  -> counts as a valid portal upload
 - Invalid                         -> does NOT count; surfaces the Invalid
                                      Comment in the Missing Portal sheet
 - Submitted also appears in a dedicated "Pending TOL Review" sheet so the
   TOL team can chase files that are waiting on review.
"""
from __future__ import annotations

import pandas as pd

from .config import (
    BILLBACK_FEE_PER_OCCURRENCE,
    BILLBACK_REASON,
    COMPLIANT,
    EXCEPTION_STATUS_EXCEPTION,
    EXCEPTION_STATUS_EXPECTED,
    MONTH_NAMES,
    PO_STATUS_CLOSED,
    PO_STATUS_PROCESSING_CODES,
    PORTAL_PENDING_STATUSES,
    PORTAL_STATUS_INVALID,
    PORTAL_VALID_STATUSES,
    SAP_FILTER_DATE_COLUMNS,
)
from .normalizer import has_value
from .supplier_exceptions import ExceptionRecord, classify_supplier


def _map_lookup(keys: pd.Series, lookup: pd.Series) -> pd.Series:
    """Map `keys` through a `lookup` Series (its index -> its values).

    Returns a value-or-missing Series aligned to `keys`. Guards the empty-lookup
    case: pandas 3.0 coerces an empty mapper to float64, which then raises
    "Cannot cast DatetimeArray to dtype float64" when the lookup holds datetimes
    (e.g. an Upload Date lookup built from zero Invalid uploads). Returning an
    all-missing Series of the lookup's own dtype keeps downstream .fillna chains
    and Excel output identical to a normal miss.
    """
    if lookup.empty:
        return pd.Series(index=keys.index, dtype=lookup.dtype)
    return keys.map(lookup)


def _apply_rows(df: pd.DataFrame, func) -> pd.Series:
    """Row-wise apply (axis=1) that is safe on empty frames.

    pandas 3.0 returns an empty *DataFrame* (mirroring the input's columns) from
    `df.apply(func, axis=1)` when the frame has no rows, which then can't be
    assigned to a single column. Return an empty object Series in that case so
    the caller's `df["col"] = _apply_rows(df, func)` always works.
    """
    if df.empty:
        return pd.Series(index=df.index, dtype=object)
    return df.apply(func, axis=1)


def _apply_series(series: pd.Series, func, dtype=object) -> pd.Series:
    """Element-wise `.apply` that is safe on an empty Series.

    pandas 3.0 returns an empty Series holding the *input's* dtype from
    `series.apply(func)` when there are no rows, rather than inferring the
    function's return type. That silently produces the wrong dtype downstream
    -- e.g. an empty "Inbound Delivery" column's `.apply(has_value)` comes back
    as string/object dtype instead of bool, and `bool_series & other_bool_series`
    then raises. Pass the expected return dtype so an empty input still yields
    a genuine typed (e.g. boolean) Series.
    """
    if series.empty:
        return pd.Series([], index=series.index, dtype=dtype)
    return series.apply(func)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------
def build_report(
    sap_df: pd.DataFrame,
    portal_df: pd.DataFrame,
    report_year: int,
    report_month: int,
    exceptions: dict[str, "ExceptionRecord"] | None = None,
    tracker_names: set[str] | None = None,
) -> dict[str, pd.DataFrame]:
    """Return a dict {sheet_name: dataframe} for every required sheet.

    `exceptions` maps normalized supplier name -> ExceptionRecord; `tracker_names`
    is every supplier the tracker knows about. Both default to empty, in which
    case every supplier is labelled "Expected to upload" and the report is
    identical to one built before this feature existed.

    Exceptions are INFORMATIONAL. They deliberately do not affect bill-back or
    the compliance percentage.
    """
    exceptions = exceptions or {}
    tracker_names = tracker_names or set()
    label = f"{MONTH_NAMES[report_month - 1]} {report_year}"

    sap = sap_df.copy()
    portal = portal_df.copy()

    # --- Scope SAP to the report month: UNION across all filter date columns.
    # A PO is in scope if ANY of Delivery Date / Appointment Date /
    # Confirmed PU Date falls in the selected month.
    per_column_masks = []
    all_blank_mask = pd.Series(True, index=sap.index)
    for col in SAP_FILTER_DATE_COLUMNS:
        dates = pd.to_datetime(sap[col], errors="coerce")
        per_column_masks.append(
            (dates.dt.year == report_year) & (dates.dt.month == report_month)
        )
        all_blank_mask &= dates.isna()
    in_month = pd.concat(per_column_masks, axis=1).any(axis=1)
    sap_excluded_count = int((~in_month).sum())
    sap_blank_date_count = int(all_blank_mask.sum())
    sap = sap.loc[in_month].copy()

    # Rows with blank PO Number cannot be matched -> exclude from matching.
    sap_valid = sap[sap["Normalized PO Number"] != ""].copy()
    portal_valid_rows = portal[portal["Normalized PO Number"] != ""].copy()

    sap_valid["Has Inbound"] = _apply_series(
        sap_valid["Inbound Delivery"], has_value, dtype=bool
    )

    # Exception Status is annotated per SAP row (not just in the summary rollup)
    # so a future change can act on it without re-plumbing the engine.
    sap_valid["Exception Status"] = [
        classify_supplier(name, number, exceptions, tracker_names)
        for name, number in zip(
            sap_valid["Vendor Name"], sap_valid["Vendor Number"], strict=True
        )
    ]

    # --- Split portal entries by File Status ---------------------------------
    portal_valid_rows["__status"] = (
        portal_valid_rows["File Status"].fillna("").astype(str).str.strip()
    )
    valid_upload = portal_valid_rows[
        portal_valid_rows["__status"].isin(PORTAL_VALID_STATUSES)
    ]
    invalid_upload = portal_valid_rows[
        portal_valid_rows["__status"] == PORTAL_STATUS_INVALID
    ]
    pending_upload = portal_valid_rows[
        portal_valid_rows["__status"].isin(PORTAL_PENDING_STATUSES)
    ]

    # Lookups keyed by normalized PO. First entry per PO wins.
    valid_lookup = (
        valid_upload
        .drop_duplicates(subset="Normalized PO Number", keep="first")
        .set_index("Normalized PO Number")[["Supplier Name", "Upload Date", "File Status"]]
    )
    invalid_lookup = (
        invalid_upload
        .drop_duplicates(subset="Normalized PO Number", keep="first")
        .set_index("Normalized PO Number")[["Supplier Name", "Upload Date", "Invalid Comment"]]
    )

    sap_pos: set[str] = set(sap_valid["Normalized PO Number"].unique())
    portal_valid_pos: set[str] = set(valid_upload["Normalized PO Number"].unique())
    portal_invalid_pos: set[str] = set(invalid_upload["Normalized PO Number"].unique())
    portal_any_pos: set[str] = set(portal_valid_rows["Normalized PO Number"].unique())

    # Annotate every SAP row with portal-side info.
    sap_valid["Portal Match"] = sap_valid["Normalized PO Number"].isin(portal_valid_pos)
    sap_valid["Portal Invalid Match"] = sap_valid["Normalized PO Number"].isin(portal_invalid_pos)
    keys = sap_valid["Normalized PO Number"]
    sap_valid["Portal Supplier Name"] = (
        _map_lookup(keys, valid_lookup["Supplier Name"])
        .fillna(_map_lookup(keys, invalid_lookup["Supplier Name"]))
        .fillna("")
    )
    sap_valid["Upload Date"] = (
        _map_lookup(keys, valid_lookup["Upload Date"])
        .fillna(_map_lookup(keys, invalid_lookup["Upload Date"]))
    )
    sap_valid["Portal File Status"] = (
        _map_lookup(keys, valid_lookup["File Status"]).fillna("")
    )
    sap_valid.loc[sap_valid["Portal Invalid Match"], "Portal File Status"] = (
        PORTAL_STATUS_INVALID
    )
    sap_valid["Invalid Reason"] = (
        _map_lookup(keys, invalid_lookup["Invalid Comment"]).fillna("")
    )

    # One row per unique PO for the per-PO sheets.
    sap_unique = sap_valid.drop_duplicates(subset="Normalized PO Number", keep="first").copy()

    # --- Compliance buckets --------------------------------------------------
    matched = sap_unique[sap_unique["Has Inbound"] & sap_unique["Portal Match"]].copy()
    matched["Compliance Status"] = COMPLIANT

    # Missing = SAP inbound exists AND no valid portal upload. Includes rows
    # where the supplier uploaded but the file was marked Invalid -- those
    # surface the rejection reason so TOL can chase the right thing.
    missing = sap_unique[
        sap_unique["Has Inbound"] & ~sap_unique["Portal Match"]
    ].copy()
    missing["Issue"] = _apply_rows(missing, _missing_issue_text)

    # Portal upload present but SAP has no inbound delivery.
    portal_no_inbound = sap_unique[
        (sap_unique["Portal Match"] | sap_unique["Portal Invalid Match"])
        & ~sap_unique["Has Inbound"]
    ].copy()
    portal_no_inbound["Issue"] = "Portal file exists but SAP has no inbound delivery."

    not_in_sap = (
        portal_valid_rows[~portal_valid_rows["Normalized PO Number"].isin(sap_pos)]
        .drop_duplicates(subset="Normalized PO Number", keep="first")
        .copy()
    )
    not_in_sap["Issue"] = "Portal PO not found in SAP export."

    closed = sap_unique[sap_unique["PO Status"] == PO_STATUS_CLOSED].copy()
    closed["Review Status"] = _apply_rows(closed, _closed_review_status)

    processing = sap_unique[
        sap_unique["PO Status"].isin(PO_STATUS_PROCESSING_CODES)
    ].copy()
    processing["Review Status"] = _apply_rows(processing, _processing_review_status)

    # POs in scope but with no inbound delivery yet.
    no_inbound_yet = sap_unique[~sap_unique["Has Inbound"]].copy()
    no_inbound_yet["Issue"] = (
        "PO is in scope for this month but no SAP inbound delivery exists yet."
    )

    # Pending TOL Review = uploads waiting for TOL approval.
    pending_unique = pending_upload.drop_duplicates(
        subset="Normalized PO Number", keep="first"
    ).copy()

    # --- Monthly Summary -----------------------------------------------------
    total_sap = len(sap_pos)
    total_portal_any = len(portal_any_pos)
    total_portal_valid = len(portal_valid_pos)
    total_portal_invalid = len(portal_invalid_pos)
    pending_count = pending_unique["Normalized PO Number"].nunique()
    total_sap_inbound = sap_unique[sap_unique["Has Inbound"]]["Normalized PO Number"].nunique()
    matched_count = matched["Normalized PO Number"].nunique()
    missing_count = missing["Normalized PO Number"].nunique()
    portal_no_sap = not_in_sap["Normalized PO Number"].nunique()
    portal_no_inbound_count = portal_no_inbound["Normalized PO Number"].nunique()
    closed_count = closed["Normalized PO Number"].nunique()
    processing_count = processing["Normalized PO Number"].nunique()
    no_inbound_count = no_inbound_yet["Normalized PO Number"].nunique()
    invalid_in_scope = sap_unique[
        sap_unique["Has Inbound"] & sap_unique["Portal Invalid Match"]
    ]["Normalized PO Number"].nunique()

    compliance_pct = (matched_count / total_sap_inbound) if total_sap_inbound else 0.0

    summary = pd.DataFrame(
        {
            "Metric": [
                "Report Month",
                "SAP Date Filter Used",
                "SAP Rows Excluded (no filter date in month)",
                "SAP Rows With All Filter Dates Blank",
                "Total Portal POs Submitted (any status)",
                "Portal POs With Valid Upload (Approved/Submitted)",
                "Portal POs Marked Invalid",
                "Portal POs Pending TOL Review",
                "Total SAP POs (in month)",
                "Total SAP POs With Inbound Delivery",
                "SAP Inbound POs With Portal File",
                "SAP Inbound POs Missing Portal File",
                "  ...of which had an Invalid upload",
                "Portal POs Not Found In SAP",
                "Portal Files With No SAP Inbound",
                "POs In Scope Without Inbound Yet",
                "Closed POs Reviewed",
                "Processing POs Reviewed",
                "Compliance Percentage",
            ],
            "Value": [
                label,
                "Union of " + ", ".join(SAP_FILTER_DATE_COLUMNS),
                sap_excluded_count,
                sap_blank_date_count,
                total_portal_any,
                total_portal_valid,
                total_portal_invalid,
                pending_count,
                total_sap,
                total_sap_inbound,
                matched_count,
                missing_count,
                invalid_in_scope,
                portal_no_sap,
                portal_no_inbound_count,
                no_inbound_count,
                closed_count,
                processing_count,
                f"{compliance_pct:.1%}",
            ],
        }
    )

    sheets = {
        "Monthly Summary": summary,
        "Portal Export Data": _portal_sheet(portal),
        "SAP Export Data": _sap_sheet(sap),
        "SAP Inbound Matched With Portal File": _matched_columns(matched),
        "SAP Inbound Missing Portal File": _missing_columns(missing),
        "Pending TOL Review": _pending_columns(pending_unique),
        "Portal File But No SAP Inbound": _portal_no_inbound_columns(portal_no_inbound),
        "Portal PO Not Found In SAP": _not_in_sap_columns(not_in_sap),
        "No Inbound Yet": _no_inbound_yet_columns(no_inbound_yet),
        "Closed POs Review": _review_columns(closed),
        "Processing POs Review": _review_columns(processing),
        "Supplier Summary": _supplier_summary(sap_unique),
        "Should Have Uploaded": _should_have_uploaded(sap_unique),
        "Exempt But Submitting": _exempt_but_submitting(sap_unique),
        "Warehouse Summary": _warehouse_summary(sap_unique),
    }
    sheets.update(_billback_sheets(missing))
    return sheets


# ---------------------------------------------------------------------------
# Row-level labels
# ---------------------------------------------------------------------------
def _missing_issue_text(row) -> str:
    if row.get("Portal Invalid Match"):
        reason = (row.get("Invalid Reason") or "").strip()
        if reason:
            return f"Portal file was marked Invalid: {reason}"
        return "Portal file was marked Invalid (no reason provided)."
    return "No portal file was submitted."


def _closed_review_status(row) -> str:
    if row["Has Inbound"] and row["Portal Match"]:
        return "Closed — Portal file present"
    if row["Has Inbound"] and not row["Portal Match"]:
        if row.get("Portal Invalid Match"):
            return "Closed — Portal upload marked Invalid"
        return "Closed — Inbound exists but no portal file (review)"
    if not row["Has Inbound"] and row["Portal Match"]:
        return "Closed — Portal file but no SAP inbound"
    return "Closed — No inbound, no portal file"


def _processing_review_status(row) -> str:
    if row["Has Inbound"] and row["Portal Match"]:
        return "Processing — Portal file present"
    if row["Has Inbound"] and not row["Portal Match"]:
        if row.get("Portal Invalid Match"):
            return "Processing — Portal upload marked Invalid"
        return "Processing — Inbound exists but no portal file"
    if not row["Has Inbound"] and row["Portal Match"]:
        return "Processing — Portal file but no SAP inbound"
    return "Processing — Pending inbound and portal file"


_ILLEGAL_SHEET_CHARS = set(r":\/?*[]")


def _billback_sheet_name(vendor_name: str, vendor_number: str, used: set) -> str:
    """Return a unique, Excel-legal bill-back sheet name (<=31 chars).

    Prefixes with 'BB-' so all bill-back tabs group together. Falls back to the
    vendor number, then 'Unknown Supplier', when the name is blank or reduces to
    nothing after illegal characters are stripped. Collisions against names
    already in `used` get a numeric suffix. Mutates `used`.
    """
    def _clean(value: str) -> str:
        stripped = "".join(
            " " if c in _ILLEGAL_SHEET_CHARS else c for c in (value or "")
        )
        return " ".join(stripped.split())  # collapse whitespace runs

    base = _clean(vendor_name) or _clean(vendor_number) or "Unknown Supplier"

    name = ("BB-" + base)[:31]
    if name in used:
        i = 2
        while True:
            suffix = f"-{i}"
            name = ("BB-" + base)[: 31 - len(suffix)] + suffix
            if name not in used:
                break
            i += 1
    used.add(name)
    return name


def _billback_supplier_tab(rows: pd.DataFrame) -> pd.DataFrame:
    """One supplier's bill-back tab: billable POs + a TOTAL row."""
    tab = pd.DataFrame(
        {
            "PO Number": rows["PO Number"].astype(str).values,
            "Warehouse": rows["Warehouse"].values,
            "PO Status": rows["PO Status"].values,
            "Appointment Date": rows["Appointment Date"].values,
            "Delivery Date": rows["Delivery Date"].values,
            "Inbound Delivery": rows["Inbound Delivery"].values,
            "Charge Reason": BILLBACK_REASON,
            "Charge (USD)": BILLBACK_FEE_PER_OCCURRENCE,
        }
    )
    n = len(tab)
    total = {c: "" for c in tab.columns}
    total["PO Number"] = f"TOTAL — {n} occurrences"
    total["Charge (USD)"] = n * BILLBACK_FEE_PER_OCCURRENCE
    return pd.concat([tab, pd.DataFrame([total])], ignore_index=True)


def _billback_sheets(missing: pd.DataFrame) -> dict:
    """Build {sheet_name: tab} for every supplier with never-uploaded POs.

    Only rows whose portal file was never submitted are billed; rows with an
    Invalid (rejected) upload are excluded. Suppliers are ordered by occurrence
    count descending so the biggest offenders' tabs come first.
    """
    if missing is None or missing.empty:
        return {}

    billable = missing
    if "Portal Invalid Match" in billable.columns:
        billable = billable[~billable["Portal Invalid Match"].fillna(False)]
    billable = billable.copy()
    if billable.empty:
        return {}

    vnum = billable["Vendor Number"].fillna("").astype(str).str.strip()
    vname = billable["Vendor Name"].fillna("").astype(str).str.strip()
    key = vnum.where(vnum != "", vname)
    key = key.where(key != "", "Unknown Supplier")
    billable = billable.assign(__vkey=key.values, __vname=vname.values)

    # Stable sort so suppliers tied on occurrence count keep a reproducible
    # (alphabetical, from groupby) order across monthly runs.
    order = (
        billable.groupby("__vkey").size()
        .sort_values(ascending=False, kind="stable")
        .index.tolist()
    )

    used: set = set()
    sheets: dict = {}
    for vkey in order:
        rows = billable[billable["__vkey"] == vkey]
        # __vkey is the vendor number, or the vendor name if no number, or
        # "Unknown Supplier"; fall back to it when no vendor name is present.
        display_name = next((n for n in rows["__vname"] if n), str(vkey))
        sheet_name = _billback_sheet_name(display_name, str(vkey), used)
        sheets[sheet_name] = _billback_supplier_tab(rows)
    return sheets


# ---------------------------------------------------------------------------
# Sheet column selectors
# ---------------------------------------------------------------------------
def _portal_sheet(portal: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "PO Number", "Supplier Name", "Upload Date", "File Status",
        "File Name", "Uploaded By", "Invalid Comment",
        "Downloaded By", "Download Date", "Normalized PO Number",
    ]
    present = [c for c in cols if c in portal.columns]
    return portal[present].copy()


def _sap_sheet(sap: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "PO Number", "Normalized PO Number", "Vendor Number", "Vendor Name",
        "Warehouse", "PO Status", "Appointment Date", "Delivery Date",
        "Confirmed PU Date", "Est PU Date",
        "Inbound Delivery", "Inbound Delivery Status",
    ]
    return sap[cols].copy()


def _matched_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "PO Number", "Vendor Number", "Vendor Name", "Warehouse", "PO Status",
        "Appointment Date", "Delivery Date", "Inbound Delivery",
        "Inbound Delivery Status", "Portal Supplier Name", "Upload Date",
        "Portal File Status", "Compliance Status",
    ]
    return df[cols].copy()


def _missing_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "PO Number", "Vendor Number", "Vendor Name", "Warehouse", "PO Status",
        "Appointment Date", "Delivery Date", "Inbound Delivery",
        "Inbound Delivery Status", "Portal File Status", "Invalid Reason", "Issue",
    ]
    return df[cols].copy()


def _pending_columns(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "PO Number": df["PO Number"],
        "Supplier Name": df["Supplier Name"],
        "Upload Date": df["Upload Date"],
        "File Name": df.get("File Name", ""),
        "Uploaded By": df.get("Uploaded By", ""),
        "File Status": df["File Status"],
        "Note": "Awaiting TOL review — counts as compliant for the supplier.",
    })


def _portal_no_inbound_columns(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "PO Number": df["PO Number"],
        "Portal Supplier Name": df["Portal Supplier Name"],
        "Upload Date": df["Upload Date"],
        "Portal File Status": df["Portal File Status"],
        "SAP Vendor Number": df["Vendor Number"],
        "SAP Vendor Name": df["Vendor Name"],
        "Warehouse": df["Warehouse"],
        "PO Status": df["PO Status"],
        "Delivery Date": df["Delivery Date"],
        "Issue": df["Issue"],
    })


def _not_in_sap_columns(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "PO Number": df["PO Number"],
        "Portal Supplier Name": df["Supplier Name"],
        "Upload Date": df["Upload Date"],
        "File Status": df.get("File Status", ""),
        "Issue": df["Issue"],
    })


def _no_inbound_yet_columns(df: pd.DataFrame) -> pd.DataFrame:
    """SAP POs in scope this month that don't have an inbound delivery yet."""
    return pd.DataFrame({
        "PO Number": df["PO Number"],
        "Vendor Number": df["Vendor Number"],
        "Vendor Name": df["Vendor Name"],
        "Warehouse": df["Warehouse"],
        "PO Status": df["PO Status"],
        "Appointment Date": df["Appointment Date"],
        "Delivery Date": df["Delivery Date"],
        "Confirmed PU Date": df["Confirmed PU Date"],
        "Est PU Date": df["Est PU Date"],
        "Portal Match": df["Portal Match"].map({True: "Yes", False: "No"}),
        "Upload Date": df["Upload Date"],
        "Issue": df["Issue"],
    })


def _review_columns(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "PO Number": df["PO Number"],
        "Vendor Number": df["Vendor Number"],
        "Vendor Name": df["Vendor Name"],
        "Warehouse": df["Warehouse"],
        "PO Status": df["PO Status"],
        "Inbound Delivery": df["Inbound Delivery"],
        "Portal Match": df["Portal Match"].map({True: "Yes", False: "No"}),
        "Portal File Status": df["Portal File Status"],
        "Upload Date": df["Upload Date"],
        "Review Status": df["Review Status"],
    })


# ---------------------------------------------------------------------------
# Group rollups
# ---------------------------------------------------------------------------
def _should_have_uploaded(sap_unique: pd.DataFrame) -> pd.DataFrame:
    """Suppliers who uploaded NOTHING despite being expected to.

    Stricter, and more damning, than "missing some": a supplier who uploaded 9 of
    10 POs has a working process with a gap; one who uploaded 0 of 10 does not
    know the process exists. Partial cases are already covered by the bill-back
    tabs, so this sheet earns its place only by isolating total failures.

    An Invalid (rejected) upload counts as "they tried" and keeps a supplier OFF
    this sheet.
    """
    columns = [
        "Vendor Number", "Vendor Name", "Exception Status",
        "Inbound POs Expected", "Portal Uploads",
    ]
    if sap_unique.empty:
        return pd.DataFrame(columns=columns)

    rows = []
    for (vendor_num, vendor_name), g in sap_unique.groupby(
        ["Vendor Number", "Vendor Name"], dropna=False
    ):
        status = g["Exception Status"].iloc[0]
        if status == EXCEPTION_STATUS_EXCEPTION:
            continue

        uploads = int(
            g[g["Portal Match"] | g["Portal Invalid Match"]][
                "Normalized PO Number"
            ].nunique()
        )
        if uploads:
            continue

        expected = int(g[g["Has Inbound"]]["Normalized PO Number"].nunique())
        if not expected:
            continue

        rows.append({
            "Vendor Number": vendor_num,
            "Vendor Name": vendor_name,
            "Exception Status": status,
            "Inbound POs Expected": expected,
            "Portal Uploads": 0,
        })

    if not rows:
        return pd.DataFrame(columns=columns)
    return (
        pd.DataFrame(rows)
        .sort_values("Inbound POs Expected", ascending=False, kind="stable")
        .reset_index(drop=True)
    )


def _exempt_but_submitting(sap_unique: pd.DataFrame) -> pd.DataFrame:
    """Exception suppliers who are uploading anyway -- a stale-exemption signal.

    A supplier excused from uploading who nonetheless submits files has, in
    practice, solved whatever problem earned them the exemption. The exemption is
    probably out of date and worth re-reviewing. (Agropur is the live example: it
    was exempted over bad batch numbers, the tracker's own "Constantly Compliant"
    sheet lists it, and it now uploads Approved files.)

    Counts an Invalid (rejected) upload as submitting -- the supplier engaged with
    the process, which is the signal we care about here.
    """
    columns = [
        "Vendor Number", "Vendor Name", "Inbound POs", "Portal Files Uploaded",
        "Of Which Rejected", "POs Still Missing A File", "Action",
    ]
    if sap_unique.empty:
        return pd.DataFrame(columns=columns)

    rows = []
    for (vendor_num, vendor_name), g in sap_unique.groupby(
        ["Vendor Number", "Vendor Name"], dropna=False
    ):
        if g["Exception Status"].iloc[0] != EXCEPTION_STATUS_EXCEPTION:
            continue

        uploaded = int(
            g[g["Portal Match"] | g["Portal Invalid Match"]][
                "Normalized PO Number"
            ].nunique()
        )
        if not uploaded:
            continue

        rejected = int(
            g[g["Portal Invalid Match"]]["Normalized PO Number"].nunique()
        )
        inbound = int(g[g["Has Inbound"]]["Normalized PO Number"].nunique())
        missing = int(
            g[g["Has Inbound"] & ~g["Portal Match"]]["Normalized PO Number"].nunique()
        )

        rows.append({
            "Vendor Number": vendor_num,
            "Vendor Name": vendor_name,
            "Inbound POs": inbound,
            "Portal Files Uploaded": uploaded,
            "Of Which Rejected": rejected,
            "POs Still Missing A File": missing,
            "Action": "Exempt but submitting — review whether the exemption is still needed.",
        })

    if not rows:
        return pd.DataFrame(columns=columns)
    return (
        pd.DataFrame(rows)
        .sort_values("Portal Files Uploaded", ascending=False, kind="stable")
        .reset_index(drop=True)
    )


def _supplier_summary(sap_unique: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (vendor_num, vendor_name), g in sap_unique.groupby(
        ["Vendor Number", "Vendor Name"], dropna=False
    ):
        total = g["Normalized PO Number"].nunique()
        with_inbound = g[g["Has Inbound"]]["Normalized PO Number"].nunique()
        found = g[g["Has Inbound"] & g["Portal Match"]]["Normalized PO Number"].nunique()
        missing = g[g["Has Inbound"] & ~g["Portal Match"]]["Normalized PO Number"].nunique()
        invalid_uploads = g[g["Has Inbound"] & g["Portal Invalid Match"]][
            "Normalized PO Number"
        ].nunique()
        portal_no_inb = g[~g["Has Inbound"] & g["Portal Match"]]["Normalized PO Number"].nunique()
        closed_n = g[g["PO Status"] == PO_STATUS_CLOSED]["Normalized PO Number"].nunique()
        processing_n = g[g["PO Status"].isin(PO_STATUS_PROCESSING_CODES)][
            "Normalized PO Number"
        ].nunique()
        status = (
            g["Exception Status"].iloc[0]
            if "Exception Status" in g.columns and len(g)
            else EXCEPTION_STATUS_EXPECTED
        )
        pct = (found / with_inbound) if with_inbound else 0.0
        rows.append({
            "Vendor Number": vendor_num,
            "Vendor Name": vendor_name,
            "Exception Status": status,
            "Total SAP POs": total,
            "SAP POs With Inbound Delivery": with_inbound,
            "Portal Files Found": found,
            "Missing Portal Files": missing,
            "Invalid Portal Uploads": invalid_uploads,
            "Portal Files With No SAP Inbound": portal_no_inb,
            "Closed POs": closed_n,
            "Processing POs": processing_n,
            "Compliance Percentage": f"{pct:.1%}",
        })
    return pd.DataFrame(rows)


def _warehouse_summary(sap_unique: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for warehouse, g in sap_unique.groupby("Warehouse", dropna=False):
        total = g["Normalized PO Number"].nunique()
        with_inbound = g[g["Has Inbound"]]["Normalized PO Number"].nunique()
        found = g[g["Has Inbound"] & g["Portal Match"]]["Normalized PO Number"].nunique()
        missing = g[g["Has Inbound"] & ~g["Portal Match"]]["Normalized PO Number"].nunique()
        invalid_uploads = g[g["Has Inbound"] & g["Portal Invalid Match"]][
            "Normalized PO Number"
        ].nunique()
        portal_no_inb = g[~g["Has Inbound"] & g["Portal Match"]]["Normalized PO Number"].nunique()
        pct = (found / with_inbound) if with_inbound else 0.0
        rows.append({
            "Warehouse": warehouse,
            "Total SAP POs": total,
            "SAP POs With Inbound Delivery": with_inbound,
            "Portal Files Found": found,
            "Missing Portal Files": missing,
            "Invalid Portal Uploads": invalid_uploads,
            "Portal Files With No SAP Inbound": portal_no_inb,
            "Compliance Percentage": f"{pct:.1%}",
        })
    return pd.DataFrame(rows)
