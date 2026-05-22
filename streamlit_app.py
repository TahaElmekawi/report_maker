import streamlit as st
import pandas as pd
from io import BytesIO

from orange import process_orange
from etisalat import process_etisalat
from vodafone import process_vodafone
from we import process_we

# -----------------------
# إعداد الصفحة + الأيقونة
# -----------------------
st.set_page_config(
    page_title="Telecom Analyzer",
    page_icon="favicon.png",  # 👈 الأيقونة
    layout="wide"
)

st.title("📊 Telecom Report Generator")

# -----------------------
# اختيار الشركة
# -----------------------
company = st.selectbox(
    "🏢 اختر الشركة",
    ["Orange", "Etisalat", "Vodafone", "WE"]
)

# -----------------------
# رفع الملف
# -----------------------
file = st.file_uploader("📥 ارفع ملف Excel", type=["xlsx"])

if file:
    st.session_state["file"] = file
    st.success("✅ تم رفع الملف")

# -----------------------
# زر التشغيل
# -----------------------
if "file" in st.session_state:

    if st.button("🚀 تشغيل التحليل"):

        file = st.session_state["file"]
        file.seek(0)  # 🔥 مهم جدًا

        output = BytesIO()

        with st.spinner("⏳ جاري المعالجة..."):

            try:

                # -----------------------
                # ORANGE
                # -----------------------
                if company == "Orange":

                    calls, imei, site, cheet = process_orange(file)

                    with pd.ExcelWriter(output, engine="openpyxl") as writer:

                        written = False

                        if isinstance(calls, pd.DataFrame) and not calls.empty:
                            calls.to_excel(writer, "calls", index=False)
                            written = True

                        if isinstance(imei, pd.DataFrame) and not imei.empty:
                            imei.to_excel(writer, "imei", index=False)
                            written = True

                        if isinstance(site, pd.DataFrame) and not site.empty:
                            site.to_excel(writer, "site", index=False)
                            written = True

                        if isinstance(cheet, pd.DataFrame) and not cheet.empty:
                            cheet.to_excel(writer, "cheet", index=False)
                            written = True

                        if not written:
                            pd.DataFrame({"Message": ["No Data Found"]}).to_excel(
                                writer, "Empty", index=False
                            )

                    name = "orange_report.xlsx"

                # -----------------------
                # ETISALAT
                # -----------------------
                elif company == "Etisalat":

                    full, calls, imei, site = process_etisalat(file)

                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        full.to_excel(writer, "Full Sheet", index=False)
                        calls.to_excel(writer, "calls_report", index=False)
                        imei.to_excel(writer, "imei_report", index=False)
                        site.to_excel(writer, "site_report", index=False)

                    name = "etisalat_report.xlsx"

                # -----------------------
                # VODAFONE
                # -----------------------
                elif company == "Vodafone":

                    data = process_vodafone(file)

                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        data["full"].to_excel(writer, "Full Sheet 📝", index=False)
                        data["tower"].to_excel(writer, "Tower Location 🌍", index=False)
                        data["linked"].to_excel(writer, "Linked Numbers 📞", index=False)
                        data["facebook"].to_excel(writer, "Facebook Profile 💻", index=False)
                        data["orders"].to_excel(writer, "Orders & Data 🧾", index=False)
                        data["service"].to_excel(writer, "Service Numbers 📩", index=False)
                        data["imei"].to_excel(writer, "IMEI Analysis 📱", index=False)

                    name = "vodafone_report.xlsx"

                # -----------------------
                # WE
                # -----------------------
                elif company == "WE":

                    data = process_we(file)

                    with pd.ExcelWriter(output, engine="openpyxl") as writer:

                        written = False

                        if "full" in data and not data["full"].empty:
                            data["full"].to_excel(writer, "Full Sheet 📝", index=False)
                            written = True

                        if "tower" in data and not data["tower"].empty:
                            data["tower"].to_excel(writer, "Tower Location 🌍", index=False)
                            written = True

                        if "linked" in data and not data["linked"].empty:
                            data["linked"].to_excel(writer, "Linked Numbers 📞", index=False)
                            written = True

                        if "facebook" in data and not data["facebook"].empty:
                            data["facebook"].to_excel(writer, "Facebook Profile 💻", index=False)
                            written = True

                        if "orders" in data and not data["orders"].empty:
                            data["orders"].to_excel(writer, "Orders & Data 🧾", index=False)
                            written = True

                        if not written:
                            pd.DataFrame({"Message": ["No Data Found"]}).to_excel(
                                writer, "Empty", index=False
                            )

                    name = "we_report.xlsx"

            except Exception as e:
                st.error(f"❌ حصل خطأ: {str(e)}")
                st.stop()

        st.success("✅ التقرير جاهز!")

        # 🔥 مهم جدًا
        output.seek(0)

        st.download_button(
            label="⬇️ تحميل التقرير",
            data=output.getvalue(),
            file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
