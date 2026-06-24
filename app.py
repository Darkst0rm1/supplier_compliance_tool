"""Entry point — defines sidebar navigation and page labels."""
from __future__ import annotations

import streamlit as st

st.set_page_config(page_title="Supplier Compliance Dashboard", layout="wide")

pg = st.navigation([
    st.Page(
        "pages/1_Supplier_Compliance_Dashboard.py",
        title="Supplier Compliance Dashboard",
    ),
    st.Page(
        "pages/2_Delivery_Fill_Rate_Dashboard.py",
        title="Delivery Fill Rate Dashboard",
    ),
    st.Page(
        "pages/3_Sales_Order_Fill_Rate_Dashboard.py",
        title="Sales Order Fill Rate Dashboard",
    ),
    st.Page(
        "pages/4_Daily_Short_Report.py",
        title="Daily Short Report Dashboard",
    ),
    st.Page(
        "pages/5_Overstock_Report.py",
        title="Overstock Report",
    ),
    st.Page(
        "pages/6_Donate_Dispose_List.py",
        title="Donate / Dispose List",
    ),
])
pg.run()
