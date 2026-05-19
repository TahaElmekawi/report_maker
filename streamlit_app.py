import streamlit as st
import pandas as pd
from io import BytesIO

from orange import process_orange
from etisalat import process_etisalat
from vodafone import process_vodafone

# -----------------------
# إعداد الصفحة
# -----------------------
st.set_page_config(
    page_title="Telecom Analyzer",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Telecom Report Generator")

# -----------------------
# اختيار الشركة
# -----------------------
company = st.selectbox(
    "🏢 اختر الشركة",
    ["Orange", "Etisalat", "Vodafone"]
)

# -----------------------
# رفع الملف
# -----------------------
file = st.file_uploader("📥 ارفع ملف Excel", type=["xlsx"])

# حفظ الملف مرة واحدة
if file:
    st.session_state["file"] = file
    st.success("✅ تم رفع الملف")

# -----------------------
# زر التشغيل
# -----------------------
if "file" in st.session_state:

    if st.button("🚀 تشغيل التحليل"):

        with st.spinner("⏳ جاري معالجة البيانات..."):

            file = st.session_state["file"]
            output = BytesIO()

            # -----------------------
            # ORANGE
            # -----------------------
            if company == "Orange":

                calls, imei, site, cheet = process_orange(file)

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    calls.to_excel(writer, sheet_name="Calls", index=False)
                    imei.to_excel(writer, sheet_name="IMEI", index=False)
                    site.to_excel(writer, sheet_name="Sites", index=False)
                    cheet.to_excel(writer, sheet_name="Full", index=False)

                file_name = "orange_report.xlsx"

            # -----------------------
            # ETISALAT
            # -----------------------
            elif company == "Etisalat":

                full, calls, imei, site = process_etisalat(file)

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    full.to_excel(writer, sheet_name="Full", index=False)
                    calls.to_excel(writer, sheet_name="Calls", index=False)
                    imei.to_excel(writer, sheet_name="IMEI", index=False)
                    site.to_excel(writer, sheet_name="Sites", index=False)

                file_name = "etisalat_report.xlsx"

            # -----------------------
            # VODAFONE
            # -----------------------
            elif company == "Vodafone":

                data = process_vodafone(file)

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    data["full"].to_excel(writer, sheet_name="Full Sheet", index=False)
                    data["tower"].to_excel(writer, sheet_name="Tower", index=False)
                    data["linked"].to_excel(writer, sheet_name="Linked", index=False)
                    data["facebook"].to_excel(writer, sheet_name="Facebook", index=False)
                    data["orders"].to_excel(writer, sheet_name="Orders", index=False)
                    data["service"].to_excel(writer, sheet_name="Service", index=False)
                    data["imei"].to_excel(writer, sheet_name="IMEI", index=False)

                file_name = "vodafone_report.xlsx"

        st.success("✅ التقرير جاهز!")

        # -----------------------
        # تحميل الملف
        # -----------------------
        st.download_button(
            label="⬇️ تحميل التقرير",
            data=output.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
