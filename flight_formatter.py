import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import io

st.title("✈️ Portal Data Formatter")

def format_time_only(raw_time):
    if pd.isna(raw_time):
        return None
    if isinstance(raw_time, str):
        try:
            return datetime.strptime(raw_time.strip(), "%H:%M").time()
        except ValueError:
            return None
    elif isinstance(raw_time, time):
        return raw_time
    return None

def extract_services(row):
    services = []
    for col in row.index:
        if isinstance(col, str) and str(row[col]).strip() == '√':
            services.append(col.strip())

    remark = str(row.get('OTHER SERVICES/REMARKS', '')).upper()
    if 'ON CALL - NEEDED ENGINEER SUPPORT' in remark:
        services.append('On Call')
    elif 'CANCELED WITHOUT NOTICE' in remark or 'CANCELLED WITHOUT NOTICE' in remark:
        services.append('Canceled without notice')
    elif 'CANCELED' in remark or 'CANCELLED' in remark:
        services.append('Cancelled Flight')
    elif 'ON CALL' in remark:
        services.append('Per Landing')

    corrected_services = []
    for service in services:
        if service == 'TECH. SUPT':
            corrected_services.append('TECH SUPPORT')
        elif service == 'HEAD SET':
            corrected_services.append('Headset')
        else:
            corrected_services.append(service)

    return ', '.join(corrected_services) if corrected_services else None

def categorize(row):
    remark = str(row.get('OTHER SERVICES/REMARKS', '')).upper()
    if 'TRANSIT' in remark:
        return '1_TRANSIT'
    elif 'ON CALL - NEEDED ENGINEER SUPPORT' in remark:
        return '2_ONCALL_ENGINEER'
    elif 'CANCELED WITHOUT NOTICE' in remark or 'CANCELLED WITHOUT NOTICE' in remark:
        return '3_CANCELED'
    elif 'ON CALL' in remark:
        return '4_ONCALL_RECORDED'
    else:
        return '5_OTHER'

def apply_cyclic_date_correction(df, time_column, date_column):
    corrected_times = []
    last_time = None
    date_offset = 0

    for idx, row in df.iterrows():
        base_date = pd.to_datetime(row[date_column]).date()
        raw_time = format_time_only(row[time_column])
        if raw_time is None:
            corrected_times.append(None)
            continue

        # If time decreases, it's a new day
        if last_time and raw_time < last_time:
            date_offset += 1
        last_time = raw_time

        corrected_datetime = datetime.combine(base_date + timedelta(days=date_offset), raw_time)
        corrected_times.append(corrected_datetime.strftime('%m/%d/%Y %H:%M'))

    return corrected_times

def process_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name='Daily Operations Report', header=4)
    df.dropna(how='all', inplace=True)
    df.rename(columns=lambda x: x.strip() if isinstance(x, str) else x, inplace=True)
    df.rename(columns={
        'REG.': 'REG',
        'TECH.\nSUPT': 'TECH. SUPT',
        'TECH. SUPT': 'TECH SUPPORT',
        'HEAD SET': 'Headset',
        'TRANSIT': 'Transit',
        'WKLY CK': 'Weekly Check',
        'DAILY CK': 'Daily Check'
    }, inplace=True)

    # Raw time formatting
    df['STA.'] = df.apply(lambda row: format_datetime(row['DATE'], row['STA']), axis=1)
    df['STD.'] = df.apply(lambda row: format_datetime(row['DATE'], row.get('STD')), axis=1)

    # Correct ATA/ATD based on time decreases
    df['ATA.'] = apply_cyclic_date_correction(df, 'ATA', 'DATE')
    df['ATD.'] = apply_cyclic_date_correction(df, 'ATD', 'DATE')

    # Fix STD/ATD if they're earlier than STA/ATA
    df['STA_dt'] = pd.to_datetime(df['STA.'], errors='coerce')
    df['STD_dt'] = pd.to_datetime(df['STD.'], errors='coerce')
    df['ATA_dt'] = pd.to_datetime(df['ATA.'], errors='coerce')
    df['ATD_dt'] = pd.to_datetime(df['ATD.'], errors='coerce')

    df.loc[df['STD_dt'] < df['STA_dt'], 'STD_dt'] += timedelta(days=1)
    df.loc[df['ATD_dt'] < df['ATA_dt'], 'ATD_dt'] += timedelta(days=1)

    df['STD.'] = df['STD_dt'].dt.strftime('%m/%d/%Y %H:%M')
    df['ATD.'] = df['ATD_dt'].dt.strftime('%m/%d/%Y %H:%M')
    df.drop(columns=['STA_dt', 'STD_dt', 'ATA_dt', 'ATD_dt'], inplace=True)

    canceled_mask = df['OTHER SERVICES/REMARKS'].str.contains('CANCELED|CANCELLED', case=False, na=False)
    df.loc[canceled_mask, 'ATA.'] = df.loc[canceled_mask, 'STA.']
    df.loc[canceled_mask, 'ATD.'] = df.loc[canceled_mask, 'STD.']

    df['Customer'] = df['FLT NO.'].astype(str).str.strip().apply(lambda x: 'XLR' if x.startswith('DHX') else x[:2])
    df['Services'] = df.apply(extract_services, axis=1)
    df['Is Canceled'] = canceled_mask
    df['Category'] = df.apply(categorize, axis=1)
    df.sort_values(by=['Category', 'STA.'], inplace=True)

    normal_rows = []
    outliers = []

    for _, row in df.iterrows():
        try:
            formatted_row = {
                'WO#': row['W/O'],
                'Station': 'KKIA',
                'Customer': row['Customer'],
                'Flight No.': row['FLT NO.'],
                'Registration Code': row['REG'],
                'Aircraft': row['A/C TYPES'],
                'Date': pd.to_datetime(row['DATE']).strftime('%m/%d/%Y'),
                'STA.': row['STA.'],
                'ATA.': row['ATA.'],
                'STD.': row['STD.'],
                'ATD.': row['ATD.'],
                'Is Canceled': row['Is Canceled'],
                'Services': row['Services'],
                'Employees': ', '.join(filter(None, [
                    str(int(row['ENGR'])) if pd.notna(row['ENGR']) and str(row['ENGR']).replace('.', '', 1).isdigit() else '',
                    str(int(row['TECH'])) if pd.notna(row['TECH']) and str(row['TECH']).replace('.', '', 1).isdigit() else ''
                ])),
                'Remarks': '',
                'Comments': ''
            }
            normal_rows.append(formatted_row)
        except Exception:
            outliers.append(row)

    final_df = pd.DataFrame(normal_rows)
    if outliers:
        final_df = pd.concat([final_df, pd.DataFrame([{}]), pd.DataFrame(outliers)], ignore_index=True)

    return final_df, df['DATE'].iloc[0] if not df.empty else None

def format_datetime(date, raw_time):
    if pd.isna(date) or pd.isna(raw_time):
        return None
    parsed_time = format_time_only(raw_time)
    if parsed_time is None:
        return None
    return datetime.combine(pd.to_datetime(date).date(), parsed_time).strftime("%m/%d/%Y %H:%M")

uploaded_file = st.file_uploader("Upload Daily Operations Report", type=["xlsx"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")
    result_df, report_date = process_file(uploaded_file)
    st.dataframe(result_df)

    if report_date:
        try:
            date_obj = pd.to_datetime(report_date)
            download_filename = date_obj.strftime("%d%b%Y").upper() + ".xlsx"
        except Exception:
            download_filename = "Formatted_Flight_Data.xlsx"
    else:
        download_filename = "Formatted_Flight_Data.xlsx"

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False)
    st.download_button("📥 Download Formatted Excel", data=output.getvalue(), file_name=download_filename)
