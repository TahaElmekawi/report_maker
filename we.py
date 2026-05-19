# -*- coding: utf-8 -*-
import pandas as pd
import re
from datetime import datetime


def process_we(file):

    # ------------------------------
    # 1. تحميل البيانات الخام
    # ------------------------------
    df_temp = pd.read_excel(file, sheet_name="Call Details For a specific Mob", header=None)

    # البحث عن صف يحتوي "ADDRES"
    header_row_idx = None
    for i, row in df_temp.iterrows():
        if row.astype(str).str.contains("ADDRES", na=False).any():
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("لم يتم العثور على صف العناوين 'ADDRES'")

    df = pd.read_excel(file, sheet_name="Call Details For a specific Mob", header=header_row_idx)

    df.columns = df.columns.str.strip()

    # ------------------------------
    # استخراج الرقم المستهدف
    # ------------------------------
    target_number = None
    for i, row in df_temp.iterrows():
        row_str = row.astype(str).str.cat(sep=' ')
        if "Details Call For Number" in row_str:
            match = re.search(r'Details Call For Number (\d+)', row_str)
            if match:
                target_number = match.group(1)
                break

    if target_number is None:
        raise ValueError("لم يتم العثور على الرقم المستهدف")

    # ------------------------------
    # معالجة البيانات (نفس الكود)
    # ------------------------------
    df["تاريخ_الاتصال"] = pd.to_datetime(df["تاريخ الاتصال"], format="%d/%m/%Y", errors="coerce")

    def parse_time(t):
        try:
            return datetime.strptime(str(t), "%H:%M:%S").time()
        except:
            return None

    df["وقت_الاتصال"] = df["وقت الاتصال"].apply(parse_time)

    def duration_to_seconds(d):
        try:
            parts = str(d).split(":")
            return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        except:
            return 0

    df["فترة_الاتصال_ثانية"] = df["فترة الاتصال"].apply(duration_to_seconds)

    target_number = str(target_number).lstrip('0')

    df["رقم الطالب"] = df["رقم الطالب"].astype(str).str.strip().str.lstrip('0')
    df["رقم المطلوب"] = df["رقم المطلوب"].astype(str).str.strip().str.lstrip('0')

    def get_call_type(row):
        if row["رقم الطالب"] == target_number:
            return "Outgoing"
        elif row["رقم المطلوب"] == target_number:
            return "Incoming"
        else:
            return "Unknown"

    df["Call_Type_Detected"] = df.apply(get_call_type, axis=1)

    # ------------------------------
    # Full Sheet
    # ------------------------------
    full_sheet = pd.DataFrame()
    full_sheet["Call Type"] = df["Call_Type_Detected"]
    full_sheet["A Number"] = target_number
    full_sheet["تاريخ الاتصال"] = df["تاريخ_الاتصال"].dt.strftime("%d/%m/%Y").fillna("")
    full_sheet["وقت الاتصال"] = df["وقت_الاتصال"]
    full_sheet["فترة الاتصال"] = df["فترة_الاتصال_ثانية"]
    full_sheet["B Number"] = df["رقم المطلوب"].astype(str).str.strip()
    full_sheet["A Number Site Address"] = df["ADDRES"]
    full_sheet["LATITUDE"] = df["LATITUDE"]
    full_sheet["LONGITUDE"] = df["LONGITUDE"]
    full_sheet["B Number Name"] = df["اسم المطلوب"]
    full_sheet["B Number Address"] = df["عنوان المطلوب"]
    full_sheet["Calling Cell ID"] = df["Calling Cell ID"]
    full_sheet["AREA"] = df["AREA"]
    full_sheet["A Number Name"] = df["اسم الطالب"]
    full_sheet["A Number Address"] = df["عنوان الطالب"]

    # ------------------------------
    # Tower
    # ------------------------------
    tower = df.groupby("ADDRES").agg(
        Count=("ADDRES", "size"),
        Latitude=("LATITUDE", "first"),
        Longitude=("LONGITUDE", "first")
    ).reset_index().rename(columns={"ADDRES": "A_Site Location GPS"})

    tower["Map"] = tower.apply(
        lambda row: f'=HYPERLINK("https://www.google.com/maps?q={row["Latitude"]},{row["Longitude"]}", "Map")'
        if pd.notna(row["Latitude"]) and pd.notna(row["Longitude"]) else '',
        axis=1
    )

    tower = tower[["Count", "A_Site Location GPS", "Latitude", "Longitude", "Map"]]
    tower = tower.sort_values("Count", ascending=False)

    # ------------------------------
    # Linked
    # ------------------------------
    linked = df.groupby("رقم المطلوب").agg(
        Calls=("رقم المطلوب", "size"),
        B_Numbers_Names=("اسم المطلوب", "first"),
        B_Addresses=("عنوان المطلوب", "first")
    ).reset_index().rename(columns={"رقم المطلوب": "Linked Numbers"})

    linked["WhatsApp 🟢"] = linked["Linked Numbers"]
    linked["Telegram 🔵"] = linked["Linked Numbers"]

    linked = linked[["Calls", "Linked Numbers", "B_Numbers_Names", "B_Addresses", "WhatsApp 🟢", "Telegram 🔵"]]

    # ------------------------------
    # Empty Sheets
    # ------------------------------
    fb_empty = pd.DataFrame(columns=[
        "Rank 🏆", "Phone Number 📱", "Full Name 👨🏻‍💼", "Facebook ID 🆔",
        "Facebook Link 🌐", "Work 🏭", "Education 🏫", "Current Location 🌍", "Hometown 🏠"
    ])

    orders_empty = pd.DataFrame(columns=[
        "Rank 🏆", "Linked Number 📞", "Orders & Data results 📑"
    ])

    # ------------------------------
    # نفس ترتيب كولاب بالظبط
    # ------------------------------
    linked = linked.sort_values("Calls", ascending=False)
    tower = tower.sort_values("Count", ascending=False)

    return {
        "full": full_sheet,
        "tower": tower,
        "linked": linked,
        "facebook": fb_empty,
        "orders": orders_empty
    }
