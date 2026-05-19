import streamlit as st
import pandas as pd
from io import BytesIO

from orange import process_orange
from etisalat import process_etisalat

st.title("📊 Telecom Report Generator")

company = st.selectbox(
    "اختر الشركة",
    ["Orange", "Etisalat"]
)

file = st.file_uploader("📥 ارفع ملف Excel", type=["xlsx"])

if file:

    with st.spinner("⏳ جاري معالجة الملف..."):

        output = BytesIO()

        if company == "Orange":

            calls, imei, site, cheet = process_orange(file)

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                calls.to_excel(writer, sheet_name="calls", index=False)
                imei.to_excel(writer, sheet_name="imei", index=False)
                site.to_excel(writer, sheet_name="site", index=False)
                cheet.to_excel(writer, sheet_name="cheet", index=False)

            file_name = "orange_report.xlsx"

        elif company == "Etisalat":

            full, calls, imei, site = process_etisalat(file)

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                full.to_excel(writer, sheet_name="Full Sheet", index=False)
                calls.to_excel(writer, sheet_name="calls_report", index=False)
                imei.to_excel(writer, sheet_name="imei_report", index=False)
                site.to_excel(writer, sheet_name="site_report", index=False)

            file_name = "etisalat_report.xlsx"

    st.success("✅ التقرير جاهز!")

    st.download_button(
        label="⬇️ تحميل التقرير",
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
