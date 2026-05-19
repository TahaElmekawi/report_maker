# -*- coding: utf-8 -*-
import pandas as pd
import re
from datetime import datetime


def process_we(file):

    # ------------------------------
    # 1. قراءة البيانات
    # ------------------------------
    df_temp = pd.read_excel(file, sheet_name=0, header=None)

    # البحث عن صف الهيدر (ADDRES)
    header_row_idx = None
    for i, row in df_temp.iterrows():
        if row.astype(str).str.contains("ADDRES", na=False).any():
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("لم يتم العثور على صف العناوين")

    df = pd.read_excel(file, sheet_name=0, header=header_row_idx)
    df.columns = df.columns.str.strip()

    # ------------------------------
    # استخراج الرقم المستهدف
    # ------------------------------
    target_number = None
    for _, row in df_temp.iterrows():
        row_str = row.astype(str).str.cat(sep=' ')
        if "Details Call For Number" in row_str:
            match = re.search(r'Details Call For Number (\d+)', row_str)
            if match:
                target_number = match.group(1)
                break

    if target_number is None:
        raise ValueError("لم يتم العثور على الرقم")

    target_number = str(target_number).lstrip('0')

    # ------------------------------
    # تنظيف البيانات
    # ------------------------------
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()

    # ------------------------------
    # معالجة الوقت والتاريخ
    # ------------------------------
    df["تاريخ_الاتصال"] = pd.to_datetime(df["تاريخ الاتصال"], errors="coerce")

    def parse_time(t):
        try:
            return datetime.strptime(str(t), "%H:%M:%S").time()
        except:
            return None

    df["وقت_الاتصال"] = df["وقت الاتصال"].apply(parse_time)

    def duration_to_seconds(d):
        try:
            h, m, s = str(d).split(":")
            return int(h)*3600 + int(m)*60 + int(s)
        except:
            return 0

    df["فترة_الاتصال_ثانية"] = df["فترة الاتصال"].apply(duration_to_seconds)

    # ------------------------------
    # توحيد الأرقام
    # ------------------------------
    df["رقم الطالب"] = df["رقم الطالب"].astype(str).str.lstrip('0')
    df["رقم المطلوب"] = df["رقم المطلوب"].astype(str).str.lstrip('0')

    # ------------------------------
    # تحديد نوع المكالمة
    # ------------------------------
    def get_call_type(row):
        if row["رقم الطالب"] == target_number:
            return "Outgoing"
        elif row["رقم المطلوب"] == target_number:
            return "Incoming"
        return "Unknown"

    df["Call_Type"] = df.apply(get_call_type, axis=1)

    # ------------------------------
    # Full Sheet
    # ------------------------------
    full_sheet = pd.DataFrame({
        "Call Type": df["Call_Type"],
        "A Number": target_number,
        "Date": df["تاريخ_الاتصال"],
        "Time": df["وقت_الاتصال"],
        "Duration": df["فترة_الاتصال_ثانية"],
        "B Number": df["رقم المطلوب"],
        "Site": df["ADDRES"],
        "LATITUDE": df["LATITUDE"],
        "LONGITUDE": df["LONGITUDE"],
        "B Name": df["اسم المطلوب"],
        "B Address": df["عنوان المطلوب"]
    })

    # ------------------------------
    # Tower
    # ------------------------------
    tower = df.groupby("ADDRES").agg(
        Count=("ADDRES", "size"),
        Latitude=("LATITUDE", "first"),
        Longitude=("LONGITUDE", "first")
    ).reset_index()

    tower["Map"] = tower.apply(
        lambda r: f'=HYPERLINK("https://www.google.com/maps?q={r["Latitude"]},{r["Longitude"]}", "Map")'
        if pd.notna(r["Latitude"]) else '',
        axis=1
    )

    # ------------------------------
    # Linked
    # ------------------------------
    linked = df.groupby("رقم المطلوب").agg(
        Calls=("رقم المطلوب", "size"),
        Name=("اسم المطلوب", "first"),
        Address=("عنوان المطلوب", "first")
    ).reset_index()

    linked["WhatsApp"] = linked["رقم المطلوب"]
    linked["Telegram"] = linked["رقم المطلوب"]

    # ------------------------------
    # Empty Sheets
    # ------------------------------
    fb = pd.DataFrame()
    orders = pd.DataFrame()

    # ------------------------------
    # Return
    # ------------------------------
    return {
        "full": full_sheet,
        "tower": tower,
        "linked": linked,
        "facebook": fb,
        "orders": orders
    }
