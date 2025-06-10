import streamlit as st
import pandas as pd
from datetime import datetime, time
import io
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.title("✈️ Portal Data Formatter (Template Export)")

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
    r = str(row.get('OTHER SERVICES/REMARKS', '')).upper()
    if 'TRANSIT' in r: return '1_TRANSIT'
    if 'ON CALL - NEEDED ENGINEER SUPPORT' in r: return '2_ONCALL_ENGINEER'
    if 'CANCELED WITHOUT NOTICE' in r or 'CANCELLED WITHOUT NOTICE' in r: return '3_CANCELED'
    if 'ON CALL' in r: return '4_ONCALL_RECORDED'
    return '5_OTHER'

def process_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name='Daily Operations Report', header=4)
    df.dropna(how='all', inplace=True)
    df.rename(columns=lambda x: x.strip() if isinstance(x, str) else x, inplace=True)
    df.rename(columns={
        'REG.': 'REG', 'TECH.\nSUPT': 'TECH. SUPT', 'TECH. SUPT': 'TECH SUPPORT',
        'HEAD SET': 'Headset', 'TRANSIT': 'Transit',
        'WKLY CK': 'Weekly Check', 'DAILY CK': 'Daily Check'
    }, inplace=True)

    df['STA.'] = df.apply(lambda r: format_datetime(r['DATE'], r['STA'], None), axis=1)
    df['ATA.'] = df.apply(lambda r: format_datetime(r['DATE'], r['ATA'], r['STA']), axis=1)
    df['STD.'] = df.apply(lambda r: format_datetime(r['DATE'], r.get('STD'), r['STA']), axis=1)
    df['ATD.'] = df.apply(lambda r: format_datetime(r['DATE'], r.get('ATD'), r['STA']), axis=1)

    cm = df['OTHER SERVICES/REMARKS'].str.contains('CANCELED|CANCELLED', case=False, na=False)
    df.loc[cm, 'ATA.'] = df.loc[cm, 'STA.']
    df.loc[cm, 'ATD.'] = df.loc[cm, 'STD.']

    df['Customer'] = df['FLT NO.'].astype(str).str.strip().str[:2].mask(df['FLT NO.'].astype(str).str.startswith('DHX'), 'XLR')
    df['Services'] = df.apply(extract_services, axis=1)
    df['Is Canceled'] = df['OTHER SERVICES/REMARKS'].str.contains('CANCELED|CANCELLED', case=False, na=False)
    df['Category'] = df.apply(categorize, axis=1)
    df.sort_values(['Category', 'STA.'], inplace=True)

    rows = []
    for _, r in df.iterrows():
        rows.append({
            'WO#': r['W/O'], 'Station': 'KKIA', 'Customer': r['Customer'],
            'Flight No.': r['FLT NO.'], 'Registration Code': r['REG'],
            'Aircraft': r['A/C TYPES'],
            'Date': r['DATE'],                # ← keep original, no time change
            'STA.': r['STA.'], 'ATA.': r['ATA.'],
            'STD.': r['STD.'], 'ATD.': r['ATD.'],
            'Is Canceled': r['Is Canceled'], 'Services': r['Services'],
            'Employees': ', '.join(filter(None, [
                str(int(r['ENGR'])) if pd.notna(r['ENGR']) and str(r['ENGR']).replace('.', '',1).isdigit() else '',
                str(int(r['TECH'])) if pd.notna(r['TECH']) and str(r['TECH']).replace('.', '',1).isdigit() else ''
            ])),
            'Remarks': '', 'Comments': ''
        })
    return pd.DataFrame(rows), df['DATE'].iloc[0] if not df.empty else None

uploaded_file = st.file_uploader("Upload Daily Operations Report", type="xlsx")
if uploaded_file:
    st.success("File uploaded!")
    result_df, report_date = process_file(uploaded_file)
    st.dataframe(result_df)

    # convert Date col to datetime.date
    result_df['Date'] = pd.to_datetime(result_df['Date']).dt.date

    # load template
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb["Template"]

    # clear prior data
    for row in ws.iter_rows(min_row=2, max_row=1000):
        for cell in row:
            cell.value = None

    # write new data
    for r_idx, row in enumerate(dataframe_to_rows(result_df, index=False, header=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            # format Date column (7th)
            if c_idx == 7:
                cell.number_format = 'MM/DD/YYYY'

    # output
    bio = io.BytesIO()
    wb.save(bio)
    fname = (pd.to_datetime(report_date).strftime("%d%b%Y").upper() + ".xlsx"
             if report_date else "output.xlsx")
    st.download_button("Download", bio.getvalue(), file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
