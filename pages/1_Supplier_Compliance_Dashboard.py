"""Supplier Documentation Compliance Report — main compliance tool."""
from __future__ import annotations

from datetime import date

import streamlit as st

from src.compliance_engine import build_report
from src.config import EXCLUDED_PO_PREFIXES, MONTH_NAMES, SAP_FILTER_DATE_COLUMNS
from src.portal_importer import PortalImportError, load_portal
from src.receiving_importer import ReceivingImportError, load_receiving
from src.report_generator import generate_workbook
from src.sap_importer import SapImportError, describe_missing_optionals, load_sap
from src.supplier_exceptions_ui import load_exceptions_or_empty, render_exception_manager


st.title("Supplier Documentation Compliance Report")
st.caption(
    "Upload SAP and Portal exports, choose a report month, then download the "
    "compliance workbook. Built for Tree of Life inbound documentation."
)

col_sap, col_portal = st.columns(2)
with col_sap:
    sap_file = st.file_uploader("1. SAP Export (.xlsx)", type=["xlsx"], key="sap")
with col_portal:
    portal_file = st.file_uploader("2. Portal Export (.xlsx)", type=["xlsx"], key="portal")

receiving_file = st.file_uploader(
    "3. Receiving Log (.xlsx) — optional",
    type=["xlsx"],
    key="receiving",
    help=(
        "The dock receiving log. Adds document accuracy — whether the batch, "
        "BBD, and quantity on the paperwork matched the goods received. This "
        "is reported alongside compliance and never changes the Compliance "
        "Percentage or the bill-back."
    ),
)

today = date.today()
years = list(range(today.year - 3, today.year + 1))
months = list(range(1, 13))

col_year, col_month = st.columns(2)
with col_year:
    sel_year = st.selectbox("Report Year", years, index=years.index(today.year))
with col_month:
    sel_month = st.selectbox(
        "Report Month",
        months,
        index=today.month - 1,
        format_func=lambda m: MONTH_NAMES[m - 1],
    )

st.caption(
    "A SAP PO is in scope for the selected month if **any** of its date "
    f"columns falls in that month: {', '.join(SAP_FILTER_DATE_COLUMNS)}. "
    "Fixed business rule — no setting to change."
)

ready = sap_file is not None and portal_file is not None

if st.button("Generate Compliance Report", type="primary", disabled=not ready):
    try:
        with st.spinner("Loading SAP file..."):
            sap_df = load_sap(sap_file)
    except SapImportError as e:
        st.error(f"SAP file error: {e}")
        st.stop()

    blank_optionals = describe_missing_optionals(sap_df)
    if blank_optionals:
        st.warning(
            "SAP export has no values for: "
            + ", ".join(blank_optionals)
            + ". Related rollups will be empty."
        )

    try:
        with st.spinner("Loading Portal file..."):
            portal_df = load_portal(portal_file, sel_year, sel_month)
    except PortalImportError as e:
        st.error(f"Portal file error: {e}")
        st.stop()

    if portal_df.empty:
        st.warning(
            f"No portal rows found for {MONTH_NAMES[sel_month - 1]} {sel_year}. "
            "The report will be generated using SAP data only."
        )

    sap_excluded = sap_df.attrs.get("excluded_po_count", 0)
    portal_excluded = portal_df.attrs.get("excluded_po_count", 0)
    if sap_excluded or portal_excluded:
        st.info(
            f"Disregarded {sap_excluded:,} SAP and {portal_excluded:,} portal PO(s) "
            f"starting with {', '.join(EXCLUDED_PO_PREFIXES)} (excluded PO type — "
            "not subject to portal documentation)."
        )

    exceptions, tracker_names, exceptions_error = load_exceptions_or_empty()
    if exceptions_error:
        st.info(exceptions_error)

    receiving_df = None
    if receiving_file is not None:
        try:
            with st.spinner("Loading Receiving Log..."):
                receiving_df = load_receiving(receiving_file, sel_year, sel_month)
        except ReceivingImportError as e:
            st.error(f"Receiving Log error: {e}")
            st.stop()

        attrs = receiving_df.attrs
        if attrs.get("skipped_sheets"):
            st.info(
                "Receiving Log: skipped sheet(s) without the audit columns — "
                + ", ".join(attrs["skipped_sheets"])
            )
        if receiving_df.empty:
            st.warning(
                f"Receiving Log has no rows for {MONTH_NAMES[sel_month - 1]} "
                f"{sel_year}. Document accuracy will not be reported."
            )
            receiving_df = None
        else:
            unmatched = attrs.get("rows_without_po", 0)
            note = (
                f"Receiving Log: {attrs.get('rows_in_month', 0):,} row(s) in "
                f"{MONTH_NAMES[sel_month - 1]} {sel_year}."
            )
            if unmatched:
                note += (
                    f" {unmatched:,} had no PO number and cannot be matched."
                )
            st.info(note)

            non_po = attrs.get("non_po_references", [])
            if non_po:
                with st.expander(
                    f"{len(non_po)} receiving-log reference(s) that aren't SAP POs"
                ):
                    st.caption(
                        "Hand-typed carrier or supplier references. They can't "
                        "be matched to SAP, so those receipts are absent from "
                        "the accuracy figures."
                    )
                    st.write(", ".join(non_po))

    with st.spinner("Applying compliance rules..."):
        sheets = build_report(
            sap_df, portal_df, sel_year, sel_month,
            exceptions=exceptions, tracker_names=tracker_names,
            receiving_df=receiving_df,
        )

    st.success("Report generated.")

    st.subheader("Monthly Summary")
    st.dataframe(sheets["Monthly Summary"], use_container_width=True, hide_index=True)

    m1, m2, m3, m4 = st.columns(4)
    summary = sheets["Monthly Summary"].set_index("Metric")["Value"]
    m1.metric("Total SAP POs", summary["Total SAP POs (in month)"])
    m2.metric("Inbound w/ Portal File", summary["SAP Inbound POs With Portal File"])
    m3.metric("Inbound Missing Portal", summary["SAP Inbound POs Missing Portal File"])
    m4.metric("Compliance %", summary["Compliance Percentage"])

    if receiving_df is not None:
        st.subheader("Document Accuracy (from Receiving Log)")
        checked = int(summary["SAP POs With A Document Accuracy Check"])
        failed = int(summary["POs Failing Any Accuracy Check"])
        disagreements = int(summary["Portal vs Receiving Log Disagreements"])
        a1, a2, a3, a4 = st.columns(4)
        a1.metric("POs Checked At Dock", checked)
        a2.metric("Failing Any Check", failed)
        a3.metric("Doc Accuracy %", summary["Document Accuracy Percentage"])
        a4.metric("Portal vs Log Conflicts", disagreements)

        total_inbound = int(summary["Total SAP POs With Inbound Delivery"])
        if total_inbound and checked < total_inbound:
            st.caption(
                f"⚠️ Coverage: only **{checked} of {total_inbound}** inbound POs "
                "were checked at the dock. Document Accuracy % describes those "
                "POs only — it is not a rate for all suppliers, and a supplier "
                "with no checks shows `n/a`. Compliance % and bill-back are "
                "unaffected by this section."
            )
        if disagreements:
            st.caption(
                f"**{disagreements} PO(s)** where the portal and the dock "
                "disagree about whether an inbound file exists — see the "
                "`Portal vs Receiving Log` sheet. Each one is either a portal "
                "data problem or a dock data-entry problem."
            )

    billback_tabs = {k: v for k, v in sheets.items() if k.startswith("BB-")}
    if billback_tabs:
        total_charge = sum(
            int(tab.iloc[-1]["Charge (USD)"]) for tab in billback_tabs.values()
        )
        st.subheader("Non-Compliant Bill-Back")
        st.caption(
            f"{len(billback_tabs)} supplier(s) billed for missing inbound "
            f"documents — total **${total_charge:,}**. One tab per supplier is "
            "included in the Excel download (sheets prefixed `BB-`)."
        )
    else:
        st.caption("No bill-back: every inbound PO had its document uploaded.")

    chase = sheets["Should Have Uploaded"]
    st.subheader("Should Have Uploaded — Nothing Received")
    if chase.empty:
        st.caption("Every supplier expected to upload submitted at least one file.")
    else:
        st.caption(
            f"**{len(chase)}** supplier(s) uploaded **nothing at all** this month "
            "despite having inbound deliveries, and are not on the exceptions list."
        )
        st.dataframe(chase, use_container_width=True, hide_index=True)

    stale = sheets["Exempt But Submitting"]
    st.subheader("Exempt But Submitting")
    if stale.empty:
        st.caption("No exempt supplier uploaded anything this month.")
    else:
        st.caption(
            f"**{len(stale)}** exempt supplier(s) uploaded files anyway. They are "
            "excused from uploading but are doing it regardless — their exemption "
            "may no longer be needed. Worth re-reviewing on the tracker."
        )
        st.dataframe(stale, use_container_width=True, hide_index=True)

    with st.spinner("Writing Excel workbook..."):
        xlsx_bytes = generate_workbook(sheets)

    file_name = f"Supplier_Compliance_{sel_year}_{sel_month:02d}.xlsx"
    st.download_button(
        "⬇️ Download Excel report",
        data=xlsx_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

render_exception_manager()
