import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

st.set_page_config(page_title="Orange Report Generator", layout="wide")

st.title("📊 Orange Excel Report Generator")

uploaded_file = st.file_uploader("📥 ارفع ملف Excel", type=["xlsx"])

if uploaded_file is not None:

    try:
        # ==============================
        # 1. قراءة الملف
        # ==============================
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        df.columns = df.columns.str.strip()

        # تنظيف
        df = df.dropna(axis=1, how='all')
        df = df.replace(r'^\s*$', pd.NA, regex=True)
        df = df.dropna(axis=1, how='all')

        # ضبط الهيدر
        df.columns = df.iloc[3]
        df = df.iloc[4:]

        # ==============================
        # التحقق من الأعمدة
        # ==============================
        required_cols = [
            'EVENT_START_TIME', 'EVENT_DIRECTION', 'OTHER_MSISDN',
            'TARGET_IMEI', 'CELL_ADDRESS'
        ]

        for col in required_cols:
            if col not in df.columns:
                st.error(f"❌ العمود {col} غير موجود")
                st.stop()

        df['EVENT_START_TIME'] = pd.to_datetime(df['EVENT_START_TIME'], errors='coerce')

        # ==============================================
        # calls
        # ==============================================
        calls_df = df[df['OTHER_MSISDN'].notna() & (df['OTHER_MSISDN'].astype(str).str.strip() != '')].copy()
        calls_df['OTHER_MSISDN'] = calls_df['OTHER_MSISDN'].astype(str)

        calls_df['is_sms'] = calls_df['EVENT_DIRECTION'].astype(str).str.contains('SMSMT', na=False, case=False)

        grouped = calls_df.groupby('OTHER_MSISDN').agg(
            Count=('OTHER_MSISDN', 'size'),
            SMS=('is_sms', 'sum'),
            B_Full_Name=('OTHER_NAME', lambda x: x.dropna().iloc[0] if x.dropna().any() else ''),
            B_Address=('OTHER_ADDRESS', lambda x: x.dropna().iloc[0] if x.dropna().any() else ''),
            B_Number_id=('OTHER_ID', lambda x: x.dropna().iloc[0] if x.dropna().any() else '')
        ).reset_index().rename(columns={'OTHER_MSISDN': 'B Number'})

        grouped = grouped.sort_values('Count', ascending=False)

        calls_sheet = grouped[['Count','B Number',  'B_Full_Name', 'B_Address', 'B_Number_id', 'SMS']]
        calls_sheet.columns = ['Count','B Number',  'B Full Name', 'B Address', 'B Number id', 'SMS']

        # ==============================================
        # imei
        # ==============================================
        imei_df = df[df['TARGET_IMEI'].notna() & (df['TARGET_IMEI'].astype(str).str.strip() != '')].copy()
        imei_df['TARGET_IMEI'] = imei_df['TARGET_IMEI'].astype(str)

        def first_last_address(group):
            group_sorted = group.sort_values('EVENT_START_TIME')
            return pd.Series({
                'First_Use_Address': group_sorted['CELL_ADDRESS'].iloc[0] if pd.notna(group_sorted['CELL_ADDRESS'].iloc[0]) else '',
                'Last_Use_Address': group_sorted['CELL_ADDRESS'].iloc[-1] if pd.notna(group_sorted['CELL_ADDRESS'].iloc[-1]) else ''
            })

        grouped_imei = imei_df.groupby('TARGET_IMEI').agg(
            Count=('TARGET_IMEI', 'size'),
            TARGET_IMEI_TYPE=('TARGET_IMEI_TYPE', lambda x: x.dropna().iloc[0] if x.dropna().any() else 'Other'),
            First_Use_Date=('EVENT_START_TIME', 'min'),
            Last_Use_Date=('EVENT_START_TIME', 'max')
        ).reset_index()

        addr = imei_df.groupby('TARGET_IMEI').apply(first_last_address).reset_index()
        grouped_imei = grouped_imei.merge(addr, on='TARGET_IMEI', how='left')

        def make_imei_link(imei):
            return f'=HYPERLINK("https://www.imei.info/calc/?imei={imei}", "Check Info")'

        grouped_imei['Device Info'] = grouped_imei['TARGET_IMEI'].apply(make_imei_link)

        imei_sheet = grouped_imei[[ 'Count', 'TARGET_IMEI', 'TARGET_IMEI_TYPE', 'Device Info',
                                  'First_Use_Date', 'Last_Use_Date', 'First_Use_Address', 'Last_Use_Address']]

        imei_sheet.columns = ['IMEI', 'Count', 'TARGET_IMEI_TYPE', 'Device Info',
                              'First_Use_Date', 'Last_Use_Date', 'First_Use_Address', 'Last_Use_Address']

        # ==============================================
        # site
        # ==============================================
        site_df = df[df['CELL_ADDRESS'].notna() & (df['CELL_ADDRESS'].astype(str).str.strip() != '')].copy()
        site_df['CELL_ADDRESS'] = site_df['CELL_ADDRESS'].astype(str)

        def first_valid_coords(group):
            lat_series = group['CELL_LAT'].dropna()
            lon_series = group['CELL_LONG'].dropna()
            return pd.Series({
                'Lat': lat_series.iloc[0] if not lat_series.empty else None,
                'Lon': lon_series.iloc[0] if not lon_series.empty else None
            })

        grouped_site = site_df.groupby('CELL_ADDRESS').agg(
            Count=('CELL_ADDRESS', 'size'),
            First_Use_Date=('EVENT_START_TIME', 'min'),
            Last_Use_Date=('EVENT_START_TIME', 'max') ,
            CELL_LAT = ("CELL_LAT"),
            CELL_LONG = ('CELL_LONG')
        ).reset_index()

        coords = site_df.groupby('CELL_ADDRESS').apply(first_valid_coords).reset_index()
        grouped_site = grouped_site.merge(coords, on='CELL_ADDRESS', how='left')

        def make_map_link(lat, lon):
            if lat and lon and pd.notna(lat) and pd.notna(lon):
                return f'=HYPERLINK("https://www.google.com/maps?q={lat},{lon}", "Map")'
            return ''

        grouped_site['Map'] = grouped_site.apply(lambda row: make_map_link(row['Lat'], row['Lon']), axis=1)

        site_sheet = grouped_site[['Count','CELL_ADDRESS',  'Map', 'CELL_LAT', 'CELL_LONG', 'First_Use_Date', 'Last_Use_Date']]
        site_sheet.columns = ['Count','CELL_ADDRESS',  'Map', 'CELL_LAT', 'CELL_LONG', 'First_Use_Date', 'Last_Use_Date']

        # ==============================================
        # cheet
        # ==============================================
        target_columns = [
            'TARGET_MSISDN', 'TARGET_IMEI', 'TARGET_IMSI', 'TARGET_IMEI_TYPE',
            'EVENT_START_TIME', 'CALL_DURATION', 'EVENT_DIRECTION', 'OTHER_MSISDN',
            'OTHER_NAME', 'OTHER_ID', 'OTHER_ID_TYPE', 'OTHER_ADDRESS',
            'CELL_ADDRESS', 'CELL_LAT', 'CELL_LONG',
            'FWD_MSISDN', 'FWD_IMEI', 'FWD_IMSI',
            'CGI', 'MCC', 'MNC', 'LAC', 'CI', 'AZMITH'
        ]

        available = [col for col in target_columns if col in df.columns]
        missing = [col for col in target_columns if col not in df.columns]

        cheet_df = df[available].copy()

        for col in missing:
            cheet_df[col] = np.nan

        cheet_df = cheet_df[target_columns]

        # ==============================================
        # Export Excel
        # ==============================================
        output_file = "orange_report.xlsx"

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            calls_sheet.to_excel(writer, sheet_name='calls', index=False)
            imei_sheet.to_excel(writer, sheet_name='imei', index=False)
            site_sheet.to_excel(writer, sheet_name='site', index=False)
            cheet_df.to_excel(writer, sheet_name='cheet', index=False)

        # ==============================
        # عرض النتائج
        # ==============================
        st.success("✅ تم إنشاء التقرير بنجاح")

        st.subheader("📞 Calls")
        st.dataframe(calls_sheet.head())

        st.subheader("📱 IMEI")
        st.dataframe(imei_sheet.head())

        st.subheader("📍 Sites")
        st.dataframe(site_sheet.head())

        # تحميل
        with open(output_file, "rb") as f:
            st.download_button(
                "⬇️ Download Report",
                f,
                file_name="orange_report.xlsx"
            )

    except Exception as e:
        st.error(f"❌ حصل خطأ: {e}")
