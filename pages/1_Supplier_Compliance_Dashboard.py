"""Supplier Documentation Compliance Report — main compliance tool."""
from __future__ import annotations

from datetime import date

import streamlit as st

from src.compliance_engine import build_report
from src.config import MONTH_NAMES, SAP_FILTER_DATE_COLUMNS
from src.portal_importer import PortalImportError, load_portal
from src.report_generator import generate_workbook
from src.sap_importer import SapImportError, describe_missing_optionals, load_sap


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

    with st.spinner("Applying compliance rules..."):
        sheets = build_report(sap_df, portal_df, sel_year, sel_month)

    st.success("Report generated.")

    st.subheader("Monthly Summary")
    st.dataframe(sheets["Monthly Summary"], use_container_width=True, hide_index=True)

    m1, m2, m3, m4 = st.columns(4)
    summary = sheets["Monthly Summary"].set_index("Metric")["Value"]
    m1.metric("Total SAP POs", summary["Total SAP POs (in month)"])
    m2.metric("Inbound w/ Portal File", summary["SAP Inbound POs With Portal File"])
    m3.metric("Inbound Missing Portal", summary["SAP Inbound POs Missing Portal File"])
    m4.metric("Compliance %", summary["Compliance Percentage"])

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

    with st.spinner("Writing Excel workbook..."):
        xlsx_bytes = generate_workbook(sheets)

    file_name = f"Supplier_Compliance_{sel_year}_{sel_month:02d}.xlsx"
    st.download_button(
        "⬇️ Download Excel report",
        data=xlsx_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
