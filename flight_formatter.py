import streamlit as st
import pandas as pd
from datetime import datetime, time
import io
from openpyxl.utils import get_column_letter

# Minimalist Title
st.markdown("""
    <h2 style='text-align: center; color: #333;'>Portal Data Formatter</h2>
    <p style='text-align: center; color: #666;'>Upload your daily operations report and get it formatted instantly.</p>
    <hr style='border: 0.5px solid #DDD;'>
""", unsafe_allow_html=True)

def format_datetime(date, raw_time):
    if pd.isna(date) or pd.isna(raw_time):
        return None
    if isinstance(raw_time, str):
        try:
            parsed_time = datetime.strptime(raw_time, "%H:%M").time()
        except ValueError:
            try:
                parsed_time = datetime.strptime(raw_time, "%H:%M:%S").time()
            except ValueError:
                return None
    elif isinstance(raw_time, time):
        parsed_time = raw_time
    else:
        return None
    parsed_time = parsed_time.replace(second=0)
    return datetime.combine(pd.to_datetime(date).date(), parsed_time).strftime("%m/%d/%Y %H:%M:%S")

def extract_services(row):
    services = []
    for col in row.index:
        if isinstance(col, str) and str(row[col]).strip() == '√':
            services.append(col.strip().title())

    remark = str(row.get('OTHER SERVICES/REMARKS', '')).upper()
    if 'ON CALL - NEEDED ENGINEER SUPPORT' in remark:
        services.append('On call - needed engineer support')
    elif 'CANCELED WITHOUT NOTICE' in remark:
        services.append('Canceled without notice')
    elif 'ON CALL' in remark:
        services.append('Per landing')

    return ', '.join(services) if services else None

def categorize(row):
    remark = str(row.get('OTHER SERVICES/REMARKS', '')).upper()
    if 'TRANSIT' in remark:
        return '1_TRANSIT'
    elif 'ON CALL - NEEDED ENGINEER SUPPORT' in remark:
        return '2_ONCALL_ENGINEER'
    elif 'CANCELED WITHOUT NOTICE' in remark:
        return '3_CANCELED'
    elif 'ON CALL' in remark:
        return '4_ONCALL_RECORDED'
    else:
        return '5_OTHER'

def process_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name='Daily Operations Report', header=4)
    df.dropna(how='all', inplace=True)
    df.rename(columns=lambda x: x.strip() if isinstance(x, str) else x, inplace=True)
    df.rename(columns={'REG.': 'REG', 'TECH.\nSUPT': 'TECH. SUPT'}, inplace=True)

    df['STA.'] = df.apply(lambda row: format_datetime(row['DATE'], row['STA']), axis=1)
    df['ATA.'] = df.apply(lambda row: format_datetime(row['DATE'], row['ATA']), axis=1)
    df['STD.'] = df.apply(lambda row: format_datetime(row['DATE'], row.get('STD')), axis=1)
    df['ATD.'] = df.apply(lambda row: format_datetime(row['DATE'], row.get('ATD')), axis=1)

    df['Customer'] = df['FLT NO.'].astype(str).str.strip().str[:2]
    df['Services'] = df.apply(extract_services, axis=1)
    df['Is Canceled'] = df['OTHER SERVICES/REMARKS'].str.contains('CANCELED', na=False, case=False)
    df['Category'] = df.apply(categorize, axis=1)
    df.sort_values(by=['Category', 'STA.'], inplace=True)

    final_df = pd.DataFrame({
        'WO#': df['W/O'],
        'Station': 'KKIA',
        'Customer': df['Customer'],
        'Flight No.': df['FLT NO.'],
        'Registration Code': df['REG'],
        'Aircraft': df['A/C TYPES'],
        'Date': pd.to_datetime(df['DATE']).dt.strftime('%m/%d/%Y'),
        'STA.': df['STA.'],
        'ATA.': df['ATA.'],
        'STD.': df['STD.'],
        'ATD.': df['ATD.'],
        'Is Canceled': df['Is Canceled'],
        'Services': df['Services'],
        'Employees': df[['ENGR', 'TECH']].apply(lambda row: str(int(row['ENGR'])) if pd.notna(row['ENGR']) and str(row['ENGR']).replace('.', '', 1).isdigit() and (pd.isna(row['TECH']) or not str(row['TECH']).replace('.', '', 1).isdigit()) else str(int(row['TECH'])) if pd.notna(row['TECH']) and str(row['TECH']).replace('.', '', 1).isdigit() and (pd.isna(row['ENGR']) or not str(row['ENGR']).replace('.', '', 1).isdigit()) else ', '.join(filter(None, [str(int(row['ENGR'])) if pd.notna(row['ENGR']) and str(row['ENGR']).replace('.', '', 1).isdigit() else '', str(int(row['TECH'])) if pd.notna(row['TECH']) and str(row['TECH']).replace('.', '', 1).isdigit() else ''])), axis=1),
        'Remarks': '',
        'Comments': ''
    })
    return final_df, df['DATE'].iloc[0] if not df.empty else None

uploaded_file = st.file_uploader("Upload Daily Operations Report", type=["xlsx"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")
    result_df, report_date = process_file(uploaded_file)
    st.dataframe(result_df, use_container_width=True)

    # Generate file name from date
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
        result_df.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.book['Sheet1']
        for i, col in enumerate(result_df.columns, 1):
            max_length = max(result_df[col].astype(str).map(len).max(), len(str(col))) + 2
            worksheet.column_dimensions[get_column_letter(i)].width = max_length

    st.download_button("📥 Download Formatted Excel", data=output.getvalue(), file_name=download_filename)



