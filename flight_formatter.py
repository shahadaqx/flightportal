import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import io

st.title("✈️ Portal Data Formatter")

def format_time(raw_time):
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

def adjust_times_with_cycle(times, base_date):
    adjusted_datetimes = []
    day_offset = 0
    previous_time = None
    for t in times:
        if pd.isna(t):
            adjusted_datetimes.append(None)
            continue
        if previous_time and t < previous_time:
            day_offset += 1
        adjusted_datetime = datetime.combine(base_date + timedelta(days=day_offset), t)
        adjusted_datetimes.append(adjusted_datetime)
        previous_time = t
    return adjusted_datetimes

def format_datetime_string(dt):
    return dt.strftime('%m/%d/%Y %H:%M') if pd.notna(dt) else None

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

    corrected = []
    for service in services:
        if service == 'TECH. SUPT':
            corrected.append('TECH SUPPORT')
        elif service == 'HEAD SET':
            corrected.append('Headset')
        else:
            corrected.append(service)
    return ', '.join(corrected) if corrected else None

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

    base_date = pd.to_datetime(df['DATE'].iloc[0]).date()

    # Convert all time columns to actual time objects
    df['STA_time'] = df['STA'].apply(format_time)
    df['ATA_time'] = df['ATA'].apply(format_time)
    df['STD_time'] = df['STD'].apply(format_time)
    df['ATD_time'] = df['ATD'].apply(format_time)

    # Adjust ATA and ATD across time cycles
    df['Adjusted_ATA'] = adjust_times_with_cycle(df['ATA_time'], base_date)
    df['Adjusted_ATD'] = adjust_times_with_cycle(df['ATD_time'], base_date)

    # STD must be >= STA
    df['STA_dt'] = [datetime.combine(base_date, t) if pd.notna(t) else None for t in df['STA_time']]
    df['STD_dt'] = [datetime.combine(base_date, t) if pd.notna(t) else None for t in df['STD_time']]
    df.loc[df['STD_dt'] < df['STA_dt'], 'STD_dt'] += timedelta(days=1)

    # Format all datetimes
    df['STA.'] = df['STA_dt'].apply(format_datetime_string)
    df['STD.'] = df['STD_dt'].apply(format_datetime_string)
    df['ATA.'] = df['Adjusted_ATA'].apply(format_datetime_string)
    df['ATD.'] = df['Adjusted_ATD'].apply(format_datetime_string)

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

# === Streamlit UI ===
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
