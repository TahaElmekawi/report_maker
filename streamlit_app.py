import streamlit as st
import pandas as pd
from io import BytesIO

from orange import process_orange
from etisalat import process_etisalat
from vodafone import process_vodafone
from we import process_we

st.set_page_config(page_title="Telecom Analyzer", layout="wide")

st.title("📊 Telecom Report Generator")

# اختيار الشركة
company = st.selectbox(
    "🏢 اختر الشركة",
    ["Orange", "Etisalat", "Vodafone", "WE"]
)

# رفع الملف
file = st.file_uploader("📥 ارفع ملف Excel", type=["xlsx"])

if file:
    st.session_state["file"] = file
    st.success("✅ تم رفع الملف")

# زر التشغيل
if "file" in st.session_state:

    if st.button("🚀 تشغيل التحليل"):

        file = st.session_state["file"]
        output = BytesIO()

        with st.spinner("⏳ جاري المعالجة..."):

            # ORANGE
            if company == "Orange":
                calls, imei, site, cheet = process_orange(file)

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    calls.to_excel(writer, "Calls", index=False)
                    imei.to_excel(writer, "IMEI", index=False)
                    site.to_excel(writer, "Sites", index=False)
                    cheet.to_excel(writer, "Full", index=False)

                name = "orange.xlsx"

            # ETISALAT
            elif company == "Etisalat":
                full, calls, imei, site = process_etisalat(file)

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    full.to_excel(writer, "Full", index=False)
                    calls.to_excel(writer, "Calls", index=False)
                    imei.to_excel(writer, "IMEI", index=False)
                    site.to_excel(writer, "Sites", index=False)

                name = "etisalat.xlsx"

            # VODAFONE
            elif company == "Vodafone":
                data = process_vodafone(file)

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    data["full"].to_excel(writer, "Full", index=False)
                    data["tower"].to_excel(writer, "Tower", index=False)
                    data["linked"].to_excel(writer, "Linked", index=False)
                    data["facebook"].to_excel(writer, "Facebook", index=False)
                    data["orders"].to_excel(writer, "Orders", index=False)
                    data["service"].to_excel(writer, "Service", index=False)
                    data["imei"].to_excel(writer, "IMEI", index=False)

                name = "vodafone.xlsx"

            # WE
            elif company == "WE":
                data = process_we(file)

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    data["full"].to_excel(writer, "Full", index=False)
                    data["tower"].to_excel(writer, "Tower", index=False)
                    data["linked"].to_excel(writer, "Linked", index=False)
                    data["facebook"].to_excel(writer, "Facebook", index=False)
                    data["orders"].to_excel(writer, "Orders", index=False)

                name = "we.xlsx"

        st.success("✅ التقرير جاهز!")

        st.download_button(
            "⬇️ تحميل التقرير",
            data=output.getvalue(),
            file_name=name
        )
