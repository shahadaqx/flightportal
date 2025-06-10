import streamlit as st
import pandas as pd
from datetime import datetime, time
import io
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Flight Formatter", layout="wide")
st.title("✈️ Portal Data Formatter (Final Version)")

TEMPLATE_FILE = "00. WorkOrdersTemplate.xlsx"

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
            else:
                return None

        raw_time_obj = to_time(raw_time)
        base_time_obj = to_time(base_time)

        if base_time_obj and base_time_obj >= time(18, 0) and raw_time_obj < time(3, 0):
            base_date += pd.Timedelta(days=1)

        return datetime.combine(base_date, raw_time_obj.replace(second=0))
    except Exception:
        return None

def extract_services(row):
    services = [col for col in row.index if isinstance(col, str) and str(row[col]).strip() == '√']
    remark = str(row.get('OTHER SERVICES/REMARKS', '')).upper()

    if 'ON CALL - NEEDED ENGINEER SUPPORT' in remark:
        services.append('On Call')
    elif 'CANCELED WITHOUT NOTICE' in remark or 'CANCELLED WITHOUT NOTICE' in remark:
        services.append('Canceled without notice')
    elif 'CANCELED' in remark or 'CANCELLED' in remark:
        services.append('Cancelled Flight')
    elif 'ON CALL' in remark:
        services.append('Per Landing')

    return ', '.join([s.replace('TECH. SUPT', 'TECH SUPPORT').replace('HEAD SET', 'Headset') for s in services])

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

    rows = []
    for _, row in df.iterrows():
        rows.append({
            'WO#': row.get('W/O'),
            'Station': 'KKIA',
            'Customer': row['Customer'],
            'Flight No.': row['FLT NO.'],
            'Registration Code': row['REG'],
            'Aircraft': row['A/C TYPES'],
            'Date': pd.to_datetime(row['DATE']).date(),  # ✅ Only date (no time)
            'STA.': row['STA.'],
            'ATA.': row['ATA.'],
            'STD.': row['STD'],
            'ATD.': row['ATD'],
            'Is Canceled': row['Is Canceled'],
            'Services': row['Services'],
            'Employees': ', '.join(filter(None, [
                str(int(row['ENGR'])) if pd.notna(row['ENGR']) and str(row['ENGR']).replace('.', '', 1).isdigit() else '',
                str(int(row['TECH'])) if pd.notna(row['TECH']) and str(row['TECH']).replace('.', '', 1).isdigit() else ''
            ])),
            'Remarks': '',
            'Comments': ''
        })

    return pd.DataFrame(rows), df['DATE'].iloc[0] if not df.empty else None

uploaded_file = st.file_uploader("📤 Upload 'Daily Operations Report'", type=["xlsx"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")
    result_df, report_date = process_file(uploaded_file)
    st.dataframe(result_df)

    try:
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb["Template"]

        # Clear existing rows
        for row in ws.iter_rows(min_row=2, max_row=1000):
            for cell in row:
                cell.value = None

        # Insert data + apply date/time formatting
        for r_idx, row in enumerate(dataframe_to_rows(result_df, index=False, header=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if isinstance(value, datetime):
                    if value.time() == datetime.min.time():
                        cell.number_format = 'MM/DD/YYYY'  # Just date
                    else:
                        cell.number_format = 'MM/DD/YYYY HH:MM:SS'  # Full datetime
                elif isinstance(value, pd.Timestamp):
                    if value.time() == datetime.min.time():
                        cell.number_format = 'MM/DD/YYYY'
                    else:
                        cell.number_format = 'MM/DD/YYYY HH:MM:SS'

        # Save to BytesIO
        output = io.BytesIO()
        wb.save(output)

        filename = (
            pd.to_datetime(report_date).strftime("%d%b%Y").upper() + ".xlsx"
            if report_date else "Formatted_WorkOrders.xlsx"
        )

        st.download_button(
            "📥 Download Final Work Orders File",
            data=output.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except FileNotFoundError:
        st.error("❌ Template file '00. WorkOrdersTemplate.xlsx' not found. Make sure it's included in your repo.")
