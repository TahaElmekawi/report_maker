# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np


def process_vodafone(file):

    # ------------------------------
    # 1. تحميل البيانات
    # ------------------------------
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()

    # الأعمدة المطلوبة
    required_cols = [
        "CALL_TYPE", "A_NUMBER", "FULL_DATE", "CALL_TIME", "SERVICE",
        "DESTINATION", "B_NUMBER", "ROUNDED_VOLUME", "SITE_ADDRESS",
        "LATITUDE", "LONGITUDE", "IMEI", "HANDSET_MANUFACTURER",
        "HANDSET_MARKETING_NAME", "SIM_SERIAL", "B_NUMBER_NATIONAL_ID",
        "B_NUMBER_SITE_ADDRESS"
    ]

    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    # ------------------------------
    # ✅ تنظيف البيانات (تم الإصلاح هنا)
    # ------------------------------
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()

    # ------------------------------
    # Full Sheet
    # ------------------------------
    full_sheet = df.copy()

    optional_cols = [
        "RATED_AMOUNT", "ALPHA_SITE_ID", "B_NUMBER_ALPHA_SITE_ID",
        "LAC", "TACCODE", "NETWORK_CARRIER", "VOWIFI_PORT",
        "VOWIFI_PUBLIC_IP"
    ]

    if all(col in df.columns for col in optional_cols):
        full_sheet = full_sheet[[
            "CALL_TYPE", "A_NUMBER", "FULL_DATE", "CALL_TIME", "SERVICE",
            "DESTINATION", "B_NUMBER", "ROUNDED_VOLUME", "RATED_AMOUNT",
            "ALPHA_SITE_ID", "SITE_ADDRESS", "LATITUDE", "LONGITUDE",
            "IMEI", "HANDSET_MANUFACTURER", "HANDSET_MARKETING_NAME",
            "SIM_SERIAL", "B_NUMBER_NATIONAL_ID", "B_NUMBER_SITE_ADDRESS",
            "B_NUMBER_ALPHA_SITE_ID", "LAC", "TACCODE",
            "NETWORK_CARRIER", "VOWIFI_PORT", "VOWIFI_PUBLIC_IP"
        ]]

    # ------------------------------
    # Tower Location
    # ------------------------------
    tower = df.groupby("SITE_ADDRESS").agg(
        Count=("SITE_ADDRESS", "size"),
        Latitude=("LATITUDE", "first"),
        Longitude=("LONGITUDE", "first")
    ).reset_index().rename(columns={"SITE_ADDRESS": "A_Site Location GPS🛰️"})

    tower["Map"] = tower.apply(
        lambda row: f'=HYPERLINK("https://www.google.com/maps?q={row["Latitude"]},{row["Longitude"]}", "Map")'
        if pd.notna(row["Latitude"]) and pd.notna(row["Longitude"]) else '',
        axis=1
    )

    tower = tower[["Count", "A_Site Location GPS🛰️", "Latitude", "Longitude", "Map"]]
    tower = tower.sort_values("Count", ascending=False)

    # ------------------------------
    # Linked Numbers
    # ------------------------------
    linked = df.groupby("B_NUMBER").agg(
        Calls=("B_NUMBER", "size"),
        Call_Duration=("ROUNDED_VOLUME", "sum"),
        B_ID_Number=("B_NUMBER_NATIONAL_ID", "first"),
        B_Site_Address=("B_NUMBER_SITE_ADDRESS", "first")
    ).reset_index().rename(columns={
        "B_NUMBER": "Linked Numbers📱",
        "B_ID_Number": "B_ID Numbers💳",
        "B_Site_Address": "B_Site Addresses🛰️"
    })

    linked = linked.rename(columns={"Call_Duration": "Call Duration⏱️"})

    linked["WhatsApp 🟢"] = linked["Linked Numbers📱"]
    linked["Telegram 🔵"] = linked["Linked Numbers📱"]

    linked = linked[[
        "Calls", "Linked Numbers📱", "Call Duration⏱️",
        "B_ID Numbers💳", "B_Site Addresses🛰️",
        "WhatsApp 🟢", "Telegram 🔵"
    ]]

    linked = linked.sort_values("Calls", ascending=False)

    # ------------------------------
    # Facebook (Empty)
    # ------------------------------
    fb_empty = pd.DataFrame(columns=[
        "Rank 🏆", "Phone Number 📱", "Other Numbers 📞",
        "Full Name 👨🏻‍💼", "Facebook ID 🆔", "Facebook Link 🌐",
        "Current Location 🌍", "Hometown 🏠",
        "Education 🏫", "Work 🏭"
    ])

    # ------------------------------
    # Orders (Empty)
    # ------------------------------
    orders_empty = pd.DataFrame(columns=[
        "Rank 🏆", "Linked Number 📞",
        "Orders & Data results 📑", "Source 💾"
    ])

    # ------------------------------
    # Service Numbers
    # ------------------------------
    service_df = df[df["DESTINATION"].notna() & (df["DESTINATION"] != "")]

    service_counts = service_df.groupby("DESTINATION").size().reset_index(name="Calls 📞")
    service_counts = service_counts.rename(columns={
        "DESTINATION": "Service Numbers & Sender ID 📩"
    })

    service_counts = service_counts.sort_values("Calls 📞", ascending=False)

    # ------------------------------
    # IMEI Analysis
    # ------------------------------
    imei_df = df[df["IMEI"] != "-1"].copy()

    imei_analysis = imei_df.groupby("IMEI").agg(
        Count=("IMEI", "size"),
        Handset_Manufacturer=("HANDSET_MANUFACTURER", "first"),
        Handset_Marketing_Name=("HANDSET_MARKETING_NAME", "first"),
        First_Use_Date=("FULL_DATE", "min"),
        Last_Use_Date=("FULL_DATE", "max"),
        First_Use_Time=("CALL_TIME", "first"),
        Last_Use_Time=("CALL_TIME", "last"),
        First_Use_Address=("SITE_ADDRESS", "first"),
        Last_Use_Address=("SITE_ADDRESS", "last")
    ).reset_index().rename(columns={
        "IMEI": "🔢 IMEI",
        "Count": "📊 Count",
        "Handset_Manufacturer": "🏭 Handset Manufacturer",
        "Handset_Marketing_Name": "🏷️ Handset Marketing Name",
        "First_Use_Date": "📅 First Use Date",
        "Last_Use_Date": "📅 Last Use Date",
        "First_Use_Time": "⏰ First Use Time",
        "Last_Use_Time": "⏰ Last Use Time",
        "First_Use_Address": "📍 First Use Address",
        "Last_Use_Address": "📍 Last Use Address"
    })

    imei_analysis["ℹ️ Device Info"] = (
        '=HYPERLINK("https://www.imei.info/?imei=" & ' +
        imei_analysis["🔢 IMEI"].astype(str) +
        ', "📲Get Info")'
    )

    imei_analysis = imei_analysis.sort_values("📊 Count", ascending=False)

    # ------------------------------
    # Return
    # ------------------------------
    return {
        "full": full_sheet,
        "tower": tower,
        "linked": linked,
        "facebook": fb_empty,
        "orders": orders_empty,
        "service": service_counts,
        "imei": imei_analysis
    }
