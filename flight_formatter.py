import streamlit as st
import pandas as pd
from datetime import datetime, time
import io
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.title("✈️ Portal Data Formatter (Preserve Template Format)")

def format_datetime(date, raw_time, base_time=None):
    if pd.isna(date) or pd.isna(raw_time):
        return None
    try:
        base_date = pd.to_datetime(date).date()
        def to_time(val):
            if isinstance(val, str):
                return datetime.strptime(val.strip(), "%H:%M").time()
            elif isinstance(val, time):
                return val
            return None
        raw_time_obj = to_time(raw_time)
        base_time_obj = to_time(base_time)
        if base_time_obj and base_time_obj >= time(18, 0) and raw_time_obj < time(3, 0):
            base_date += pd.Timedelta(days=1)
        full_datetime = datetime.combine(base_date, raw_time_obj.replace(second=0))
        return full_datetime.strftime("%m/%d/%Y %H:%M:%S")
    except Exception:
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

    df['STA.'] = df.apply(lambda row: format_datetime(row['DATE'], row['STA'], None), axis=1)
    df['ATA.'] = df.apply(lambda row: format_datetime(row['DATE'], row['ATA'], row['STA']), axis=1)
    df['STD.'] = df.apply(lambda row: format_datetime(row['DATE'], row.get('STD'), row['STA']), axis=1)
    df['ATD.'] = df.apply(lambda row: format_datetime(row['DATE'], row.get('ATD'), row['STA']), axis=1)

    canceled_mask = df['OTHER SERVICES/REMARKS'].str.contains('CANCELED|CANCELLED', case=False, na=False)
    df.loc[canceled_mask, 'ATA.'] = df.loc[canceled_mask, 'STA.']
    df.loc[canceled_mask, 'ATD.'] = df.loc[canceled_mask, 'STD.']

    df['Customer'] = df['FLT NO.'].astype(str).str.strip().apply(lambda x: 'XLR' if x.startswith('DHX') else x[:2])
    df['Services'] = df.apply(extract_services, axis=1)
    df['Is Canceled'] = df['OTHER SERVICES/REMARKS'].str.contains('CANCELED|CANCELLED', na=False, case=False)
    df['Category'] = df.apply(categorize, axis=1)
    df.sort_values(by=['Category', 'STA.'], inplace=True)

    result_rows = []
    for _, row in df.iterrows():
        try:
            result_rows.append({
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
            })
        except:
            continue

    return pd.DataFrame(result_rows), df['DATE'].iloc[0] if not df.empty else None

uploaded_file = st.file_uploader("📤 Upload Daily Operations Report", type=["xlsx"])
template_uploaded = st.file_uploader("📥 Upload Excel Template (with formatting)", type=["xlsx"])

if uploaded_file and template_uploaded:
    st.success("✅ Files uploaded successfully.")
    result_df, report_date = process_file(uploaded_file)
    st.dataframe(result_df)

    # Load template in memory
    template_io = io.BytesIO(template_uploaded.read())
    wb = load_workbook(template_io)
    ws = wb.worksheets[0]

    # Find header row (first non-empty row)
    header_row_idx = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=20), start=1):
        if any(cell.value for cell in row):
            header_row_idx = i
            break
    if not header_row_idx:
        st.error("❌ Could not find a header row in template.")
        st.stop()
    start_row = header_row_idx + 1

    # Write values only (preserve formatting)
    for r_idx, row in enumerate(dataframe_to_rows(result_df, index=False, header=False), start=start_row):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = value  # value only; do not touch styles

    # Save result to new downloadable file
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    output_name = (
        pd.to_datetime(report_date).strftime("%d%b%Y").upper() + "_WorkOrders.xlsx"
        if report_date else "Formatted_WorkOrders.xlsx"
    )
    st.download_button("📥 Download Final Work Orders File", data=output, file_name=output_name)
