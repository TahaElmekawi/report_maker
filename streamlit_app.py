import streamlit as st
import pandas as pd

from orange import process_orange
from etisalat import process_etisalat

st.title("📊 Telecom Report Generator")

company = st.selectbox(
    "اختر الشركة",
    ["Orange", "Etisalat"]
)

file = st.file_uploader("📥 ارفع ملف Excel", type=["xlsx"])

if file:

    if company == "Orange":

        calls, imei, site, cheet = process_orange(file)

        with pd.ExcelWriter("orange.xlsx", engine="openpyxl") as writer:
            calls.to_excel(writer, sheet_name="calls", index=False)
            imei.to_excel(writer, sheet_name="imei", index=False)
            site.to_excel(writer, sheet_name="site", index=False)
            cheet.to_excel(writer, sheet_name="cheet", index=False)

        with open("orange.xlsx", "rb") as f:
            st.download_button("⬇️ Download Orange Report", f)

    elif company == "Etisalat":

        full, calls, imei, site = process_etisalat(file)

        with pd.ExcelWriter("etisalat.xlsx", engine="openpyxl") as writer:
            full.to_excel(writer, sheet_name="Full Sheet", index=False)
            calls.to_excel(writer, sheet_name="calls_report", index=False)
            imei.to_excel(writer, sheet_name="imei_report", index=False)
            site.to_excel(writer, sheet_name="site_report", index=False)

        with open("etisalat.xlsx", "rb") as f:
            st.download_button("⬇️ Download Etisalat Report", f)
