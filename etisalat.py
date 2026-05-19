def process_etisalat(file):
    import pandas as pd
    import numpy as np

    df = pd.read_excel(file, sheet_name='cheet', dtype=str)

    # تحويل البيانات
    df['Call_Start_Date'] = pd.to_datetime(df['Call_Start_Date'], errors='coerce')
    df['Actual_Duration'] = pd.to_numeric(df['Actual_Duration'], errors='coerce').fillna(0)

    # =========================================
    # calls
    # =========================================
    calls_report = []

    for b_number, group in df.groupby('B_Number', dropna=False):
        if pd.isna(b_number) or b_number == '':
            continue

        count = len(group)
        sms_count = group[group['Network_Activity_Type_Name'] == 'SMS'].shape[0]

        full_name = group['B_Number_Full_Name'].dropna().iloc[0] if not group['B_Number_Full_Name'].dropna().empty else ''
        address = group['B_Number_Address'].dropna().iloc[0] if not group['B_Number_Address'].dropna().empty else ''
        site_address = group['B_Number_MU_Site_Address'].dropna().iloc[0] if not group['B_Number_MU_Site_Address'].dropna().empty else ''
        lat = group['B_Number_MU_Latitude'].dropna().iloc[0] if not group['B_Number_MU_Latitude'].dropna().empty else ''
        lon = group['B_Number_MU_Longitude'].dropna().iloc[0] if not group['B_Number_MU_Longitude'].dropna().empty else ''

        map_link = f'=HYPERLINK("https://www.google.com/maps?q={lat},{lon}", "Map")' if lat and lon else ''

        calls_report.append({
            'B Number': b_number,
            'Count': count,
            'B Full Name': full_name,
            'B Address': address,
            'B_NUMBER_SITE_ADDRESS': site_address,
            'Latitude': lat,
            'Longitude': lon,
            'Map': map_link,
            'SMS': sms_count
        })

    df_calls = pd.DataFrame(calls_report).sort_values('Count', ascending=False)

    # =========================================
    # imei
    # =========================================
    imei_report = []

    for imei, group in df.groupby('IMEI_Number', dropna=False):
        if pd.isna(imei) or imei == '':
            continue

        count = len(group)
        first_use = group['Call_Start_Date'].min()
        last_use = group['Call_Start_Date'].max()

        if first_use is not pd.NaT and last_use is not pd.NaT:
            first_row = group.loc[group['Call_Start_Date'].idxmin()]
            last_row = group.loc[group['Call_Start_Date'].idxmax()]
            first_addr = first_row.get('Site_Address', '')
            last_addr = last_row.get('Site_Address', '')
        else:
            first_addr = last_addr = ''

        hyperlink = f'=HYPERLINK("https://www.imei.info/calc/?imei={imei}", "Device Info")'

        imei_report.append({
            'IMEI': imei,
            'Count': count,
            'Device Info': hyperlink,
            'First_Use_Date': first_use,
            'Last_Use_Date': last_use,
            'First_Use_Address': first_addr,
            'Last_Use_Address': last_addr
        })

    df_imei = pd.DataFrame(imei_report).sort_values('Count', ascending=False)

    # =========================================
    # site
    # =========================================
    site_report = []

    for site, group in df.groupby('Site_Address', dropna=False):
        if pd.isna(site) or site == '':
            continue

        count = len(group)
        first_use = group['Call_Start_Date'].min()
        last_use = group['Call_Start_Date'].max()

        lat = group['Latitude'].dropna().iloc[0] if not group['Latitude'].dropna().empty else ''
        lon = group['Longitude'].dropna().iloc[0] if not group['Longitude'].dropna().empty else ''

        map_link = f'=HYPERLINK("https://www.google.com/maps?q={lat},{lon}", "Map")' if lat and lon else ''

        site_report.append({
            'Site_Address': site,
            'Count': count,
            'Latitude': lat,
            'Longitude': lon,
            'Map': map_link,
            'First_Use_Date': first_use,
            'Last_Use_Date': last_use
        })

    df_site = pd.DataFrame(site_report).sort_values('Count', ascending=False)

    return df, df_calls, df_imei, df_site
