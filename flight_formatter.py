import streamlit as st
import pandas as pd
from datetime import datetime, time
import io
from openpyxl import load_workbook
from copy import copy

st.title("✈️ Portal Data Formatter (with Final Rollover Logic)")

# Final smart rollover logic — STA late night & raw_time early = next day
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

        if base_time_obj:
            if base_time_obj >= time(18, 0) and raw_time_obj < time(3, 0):
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

def merge_with_template(result_df, template_path):
    """Merge the processed data with the template while preserving formatting"""
    # Load the template workbook
    wb = load_workbook(template_path)
    ws = wb['Template']
    
    # Clear existing data below headers (row 2 and below)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.value = None
    
    # Write new data starting from row 2
    for r_idx, row in enumerate(result_df.itertuples(), start=2):
        for c_idx, value in enumerate(row[1:], start=1):  # Skip index (0)
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            
            # Copy formatting from header row
            if c_idx <= ws.max_column:
                header_cell = ws.cell(row=1, column=c_idx)
                cell.font = copy(header_cell.font)
                cell.border = copy(header_cell.border)
                cell.fill = copy(header_cell.fill)
                cell.number_format = copy(header_cell.number_format)
                cell.alignment = copy(header_cell.alignment)
    
    # Save to bytes buffer
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Load template from embedded data
template_data = io.BytesIO()
with pd.ExcelWriter(template_data, engine='openpyxl') as writer:
    # Create Lookups sheet
    lookups_df = pd.DataFrame({
        'A': ['KKIA'] + [''] * 70,
        'B': ['2S', '2W', '4M', '5K', '5O', '5W', '6E', '6K', '6Y', '8U', '9H', '9P', 'AFG', 'AH', 'AI', 'AT', 
              'AZ', 'B4', 'BI', 'BJ', 'BS', 'C6', 'D3', 'E5', 'EK', 'ER', 'ET', 'EY', 'F3', 'FZ', 'G9', 'GA', 
              'H9', 'IA', 'IR', 'IX', 'IY', 'J4', 'JT', 'KU', 'LN', 'LO', 'LY', 'ME', 'NB', 'PA', 'PC', 'PF', 
              'PK', 'QG', 'QP', 'QR', 'R5', 'RA', 'RJ', 'RQ', 'SG', 'SQ', 'SV', 'T7', 'T8', 'TK', 'TR', 'UK', 
              'UL', 'UZS', 'W4', 'WY', 'XLR', 'XY', 'YL', 'YR'],
        'C': ['A300', 'A310', 'A319', 'A320', 'A320 NOE', 'A321', 'A321 NOE', 'A330', 'A330 NOE', 'A340', 'A350', 
              'A380', 'B190', 'B707', 'B717', 'B720', 'B727', 'B737', 'B737 MAX', 'B737 NG', 'B747', 'B757', 'B767', 
              'B777', 'B787', 'DC10', 'DC3', 'DC6', 'DC8', 'DC9', 'E110', 'E120', 'E135', 'E145', 'E170', 'E175', 
              'E190', 'E195', 'E75S', 'EMB 140', 'EMB 170', 'EMB 190', 'MD11', 'MD80', 'MD81', 'MD82', 'MD83', 
              'MD87', 'MD88', 'MD90'] + [''] * 22
    })
    lookups_df.to_excel(writer, sheet_name='Lookups', index=False)
    
    # Create Template sheet
    template_headers = ['WO#', 'Station', 'Customer', 'Flight No.', 'Registration Code', 'Aircraft', 'Date', 'STA.', 
                       'ATA.', 'STD.', 'ATD.', 'Is Canceled', 'Services', 'Employees', 'Remarks', 'Comments']
    template_df = pd.DataFrame(columns=template_headers)
    template_df.to_excel(writer, sheet_name='Template', index=False)

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

    # Create template in memory
    template_data.seek(0)
    
    # Merge with template
    merged_output = merge_with_template(result_df, template_data)
    
    st.download_button(
        "📥 Download Formatted Excel", 
        data=merged_output.getvalue(), 
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
