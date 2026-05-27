"""Generate the two blank Excel templates under ./templates.

The SAP template mirrors the real SAP export column names (Plant, Vendor,
Appt. Date, etc.). The portal template mirrors the real Inbound Delivery
File List columns seen on the employee portal.

Run once after install:

    python make_templates.py
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd


TEMPLATES_DIR = Path(__file__).parent / "templates"

# These mirror the column names found in real SAP exports.
SAP_TEMPLATE_COLUMNS = [
    "PO Number",
    "Confirmed PU Date",
    "Est PU Date",
    "Appt. Date",
    "Delivery Date",
    "Inbound Delivery Status",
    "Inbound Delivery",
    "Vendor Name",
    "Plant",
    "Vendor",
]

# Mirrors the visible columns of the Inbound Delivery File List portal page.
PORTAL_TEMPLATE_COLUMNS = [
    "PO Number",
    "File Name",
    "Uploaded By",
    "Supplier Name",
    "File Status",
    "Upload Date",
    "Downloaded By",
    "Download Date",
    "Invalid Comment",
]


def main() -> None:
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)

    sap_sample = pd.DataFrame(
        [
            {
                "PO Number": "1177377", "Confirmed PU Date": "", "Est PU Date": "",
                "Appt. Date": "2026-05-06", "Delivery Date": "2026-05-06",
                "Inbound Delivery Status": "C", "Inbound Delivery": "0180000084",
                "Vendor Name": "EPICUREAN INTL-THAI KIT COC MK",
                "Plant": "2910", "Vendor": "70007031",
            },
            {
                "PO Number": "1000004749",
                "Confirmed PU Date": "2026-05-20", "Est PU Date": "2026-05-20",
                "Appt. Date": "", "Delivery Date": "2026-06-19",
                "Inbound Delivery Status": "A", "Inbound Delivery": "0180012333",
                "Vendor Name": "BROTHERS DRINKS CO. LTD",
                "Plant": "2910", "Vendor": "70007050",
            },
            {
                "PO Number": "1000004751", "Confirmed PU Date": "",
                "Est PU Date": "2026-05-14", "Appt. Date": "",
                "Delivery Date": "2026-06-16",
                "Inbound Delivery Status": "B", "Inbound Delivery": "0180012334",
                "Vendor Name": "BROTHERS DRINKS CO. LTD",
                "Plant": "2910", "Vendor": "70007050",
            },
        ],
        columns=SAP_TEMPLATE_COLUMNS,
    )

    portal_sample = pd.DataFrame(
        [
            {
                "PO Number": "1177377", "File Name": "PO 1177377.xlsx",
                "Uploaded By": "fcko@kikkoman.com",
                "Supplier Name": "EPICUREAN INTL-THAI KIT COC MK",
                "File Status": "Approved", "Upload Date": "2026-05-04 10:15",
                "Downloaded By": "linda.vlasblom@treeoflife.com",
                "Download Date": "2026-05-05 09:00", "Invalid Comment": "",
            },
            {
                "PO Number": "1000004749", "File Name": "BrothersDrinks_4749.xlsx",
                "Uploaded By": "ops@brothersdrinks.com",
                "Supplier Name": "BROTHERS DRINKS CO. LTD",
                "File Status": "Submitted", "Upload Date": "2026-05-18 14:32",
                "Downloaded By": "", "Download Date": "", "Invalid Comment": "",
            },
            {
                "PO Number": "1000004751", "File Name": "BrothersDrinks_4751.xlsx",
                "Uploaded By": "ops@brothersdrinks.com",
                "Supplier Name": "BROTHERS DRINKS CO. LTD",
                "File Status": "Invalid", "Upload Date": "2026-05-20 09:11",
                "Downloaded By": "chandrakala.bisht@treeoflife.com",
                "Download Date": "2026-05-21 08:00",
                "Invalid Comment": "Missing INVR signature page",
            },
        ],
        columns=PORTAL_TEMPLATE_COLUMNS,
    )

    sap_path = TEMPLATES_DIR / "sap_template.xlsx"
    portal_path = TEMPLATES_DIR / "portal_template.xlsx"
    sap_sample.to_excel(sap_path, index=False)
    portal_sample.to_excel(portal_path, index=False)
    print(f"Wrote {sap_path}")
    print(f"Wrote {portal_path}")


if __name__ == "__main__":
    main()
