
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

def categorize_services(row):
    if row['Tech Support'] == "√":
        return "ON CALL - NEEDED ENGINEER SUPPORT", "2_TECH_SUPPORT"
    elif "CANCELED" in str(row['Remarks']).upper():
        return "CANCELED", "3_CANCELED"
    elif row['Transit'] == "√":
        services = ["Transit"]
        if row['Headset'] == "√":
            services.append("Headset")
        if row['Daily Check'] == "√":
            services.append("Daily Check")
        if row['Weekly Check'] == "√":
            services.append("Weekly Check")
        return ", ".join(services), "1_TRANSIT"
    else:
        return "Per Landing", "4_ON_CALL"

def format_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=0, skiprows=4)

    df['Services'], df['SortKey'] = zip(*df.apply(categorize_services, axis=1))
    df['WO#'] = df['W/O']
    df['Station'] = "KKIA"
    df['Customer'] = df['Flight No.'].str[:2]
    df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%m/%d/%Y')
    df['STA.'] = pd.to_datetime(df['STA']).dt.strftime('%m/%d/%Y %H:%M:%S')
    df['ATA.'] = pd.to_datetime(df['ATA']).dt.strftime('%m/%d/%Y %H:%M:%S')
    df['STD.'] = pd.to_datetime(df['STD']).dt.strftime('%m/%d/%Y %H:%M:%S')
    df['ATD.'] = pd.to_datetime(df['ATD']).dt.strftime('%m/%d/%Y %H:%M:%S')
    df['Is Canceled'] = df['Services'].str.upper().str.contains("CANCELED")
    df['Employees'] = df[['EMP1', 'EMP2']].fillna('').agg(', '.join, axis=1).str.strip(', ')
    df['Remarks'] = ""
    df['Comments'] = ""

    output_columns = ['WO#', 'Station', 'Customer', 'Flight No.', 'Registration', 'Aircraft', 'Date',
                      'STA.', 'ATA.', 'STD.', 'ATD.', 'Is Canceled', 'Services', 'Employees', 'Remarks', 'Comments', 'SortKey']

    df = df[output_columns]
    df = df.sort_values(by='SortKey')
    df.drop(columns='SortKey', inplace=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Template', index=False)
    output.seek(0)
    return output

st.title("Daily Ops Formatter")
uploaded_file = st.file_uploader("Upload Daily Ops Excel file", type=["xlsx"])

if uploaded_file:
    result = format_excel(uploaded_file)
    st.download_button(
        label="Download Processed Excel",
        data=result,
        file_name="Formatted_Daily_Ops.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
