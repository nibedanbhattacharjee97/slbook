import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import base64
from io import BytesIO

# Function to create the `bani` table in `slide.db`
def create_bani_table():
    conn = sqlite3.connect('slide.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS bani (
                 id INTEGER PRIMARY KEY AUTOINCREMENT,
                 cmis_id TEXT,
                 student_name TEXT,
                 cmis_ph_no TEXT,
                 center_name TEXT,
                 uploader_name TEXT,
                 verification_type TEXT,
                 mode_of_verification TEXT,
                 date_of_verification TEXT
                 )''')
    conn.commit()
    conn.close()

# Function to validate and insert data into the database
def update_another_database(file):
    try:
        df = pd.read_excel(file)
        
        # Ensure column names match
        required_columns = ['CMIS ID', 'Student Name', 'CMIS PH No(10 Number)',
                            'Center Name', 'Name Of Uploder', 'Verification Type',
                            'Mode Of Verification', 'Date Of Verification']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Uploaded file is missing one or more required columns: {required_columns}")
            return
        
        # Rename columns to match database schema
        df.rename(columns={
            'CMIS ID': 'cmis_id',
            'Student Name': 'student_name',
            'CMIS PH No(10 Number)': 'cmis_ph_no',
            'Center Name': 'center_name',
            'Name Of Uploder': 'uploader_name',
            'Verification Type': 'verification_type',
            'Mode Of Verification': 'mode_of_verification',
            'Date Of Verification': 'date_of_verification'
        }, inplace=True)
        
        # Connect to the database
        conn = sqlite3.connect('slide.db')
        c = conn.cursor()
        
        # Insert data into the database
        for _, row in df.iterrows():
            try:
                c.execute('''INSERT INTO bani (cmis_id, student_name, cmis_ph_no, center_name, 
                             uploader_name, verification_type, mode_of_verification, date_of_verification)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?)''', 
                          (row['cmis_id'], row['student_name'], row['cmis_ph_no'], row['center_name'],
                           row['uploader_name'], row['verification_type'], row['mode_of_verification'], 
                           row['date_of_verification']))
            except sqlite3.IntegrityError as e:
                st.warning(f"Skipping record due to error: {e}")
        
        conn.commit()
        conn.close()
        st.success('Data updated successfully!')
    except Exception as e:
        st.error(f"An error occurred: {e}")

# Function to download the data from the `bani` table
def download_another_database_data():
    conn = sqlite3.connect('slide.db')
    df = pd.read_sql_query("SELECT * FROM bani", conn)
    conn.close()
    
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="bani.csv">Download CSV</a>'
    st.markdown(href, unsafe_allow_html=True)

# Function to generate and download a sample Excel file
def download_sample_excel():
    sample_data = {
        'CMIS ID': ['123', '456', '789'],
        'Student Name': ['John Doe', 'Jane Smith', 'Jim Beam'],
        'CMIS PH No(10 Number)': ['1234567890', '0987654321', '1122334455'],
        'Center Name': ['Center 1', 'Center 2', 'Center 3'],
        'Name Of Uploder': ['Uploader 1', 'Uploader 2', 'Uploader 3'],
        'Verification Type': ['Placement', 'Placement', 'Enrollment'],
        'Mode Of Verification': ['G-meet', 'Call', 'Call'],
        'Date Of Verification': ['2024-01-12', '2024-01-12', '2024-01-12']
    }

    df = pd.DataFrame(sample_data)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)
    output.seek(0)

    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="Sample_Excel.xlsx">Download Sample Excel</a>'
    st.markdown(href, unsafe_allow_html=True)

# Main function for the Streamlit app
def main():
    st.title('Student Data Upload Platform')

    # Ensure the `bani` table exists
    create_bani_table()

    # Upload Excel file and update the database
    st.subheader('Upload Student Data')
    file = st.file_uploader('Upload Excel', type=['xlsx', 'xls'])
    if file is not None:
        if st.button('Update Data'):
            update_another_database(file)

    # Download sample Excel file
    st.subheader('Download Sample Excel File')
    if st.button('Download Sample'):
        download_sample_excel()

    # Download data button
    st.subheader('Download Database Data')
    if st.button('Download Data'):
        download_another_database_data()

# Run the app
if __name__ == '__main__':
    main()
