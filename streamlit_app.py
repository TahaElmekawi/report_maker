import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Excel Report Generator", layout="wide")

st.title("📊 Excel Report Generator")

# رفع الملف
uploaded_file = st.file_uploader("ارفع ملف Excel", type=["xlsx"])

if uploaded_file is not None:

    # قراءة الملف
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    st.subheader("📌 Preview Data")
    st.dataframe(df.head())

    # ------------------------------
    # تنظيف البيانات
    # ------------------------------
    df = df.dropna(axis=1, how='all')
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    df = df.dropna(axis=1, how='all')

    # ضبط الهيدر
    df.columns = df.iloc[3]
    df = df.iloc[4:]

    # ------------------------------
    # التأكد من الأعمدة المطلوبة
    # ------------------------------
    required_cols = [
        'EVENT_START_TIME',
        'EVENT_DIRECTION',
        'OTHER_MSISDN',
        'TARGET_IMEI',
        'CELL_ADDRESS'
    ]

    missing = [col for col in required_cols if col not in df.columns]

    if missing:
        st.error(f"❌ الأعمدة دي مش موجودة: {missing}")
    else:

        # تحويل الوقت
        df['EVENT_START_TIME'] = pd.to_datetime(df['EVENT_START_TIME'], errors='coerce')

        # ------------------------------
        # تحليل calls
        # ------------------------------
        calls_df = df[
            df['OTHER_MSISDN'].notna() &
            (df['OTHER_MSISDN'].astype(str).str.strip() != '')
        ].copy()

        calls_df['OTHER_MSISDN'] = calls_df['OTHER_MSISDN'].astype(str)

        calls_df['is_sms'] = calls_df['EVENT_DIRECTION'].astype(str).str.contains(
            'SMSMT', case=False, na=False
        )

        grouped = calls_df.groupby('OTHER_MSISDN').agg(
            Count=('OTHER_MSISDN', 'size'),
            SMS=('is_sms', 'sum')
        ).reset_index()

        st.subheader("📞 Calls Analysis")
        st.dataframe(grouped)

        # ------------------------------
        # تحميل النتيجة
        # ------------------------------
        output_file = "output_report.xlsx"
        grouped.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="⬇️ Download Report",
                data=f,
                file_name="report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
