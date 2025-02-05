import streamlit as st
import pandas as pd
import sqlite3
import base64
from io import BytesIO
from datetime import datetime

# Database File Paths
STUDENT_DB = 'duplicate.db'
SLOT_BOOKING_DB = 'slot_booking_new.db'

# Helper Functions for Database Operations
def create_databases():
    """Create tables for both databases if not exist."""
    # Student Data Table
    conn = sqlite3.connect(STUDENT_DB)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS studentcap (
                 id INTEGER PRIMARY KEY AUTOINCREMENT,
                 cmis_id TEXT,
                 student_name TEXT,
                 cmis_ph_no TEXT,
                 center_name TEXT,
                 uploader_name TEXT,
                 verification_type TEXT,
                 mode_of_verification TEXT)''')
    conn.commit()
    conn.close()

    # Slot Booking Table
    conn = sqlite3.connect(SLOT_BOOKING_DB)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS appointment_bookings (
                 id INTEGER PRIMARY KEY AUTOINCREMENT,
                 date TEXT,
                 time_range TEXT,
                 manager TEXT,
                 spoc TEXT,
                 booked_by TEXT)''')
    conn.commit()
    conn.close()

def update_student_database(file):
    """Update student database from uploaded Excel file."""
    df = pd.read_excel(file)

    conn = sqlite3.connect(STUDENT_DB)
    c = conn.cursor()
    for index, row in df.iterrows():
        c.execute('''INSERT INTO studentcap (cmis_id, student_name, cmis_ph_no, center_name, uploader_name, verification_type, mode_of_verification)
                     VALUES (?, ?, ?, ?, ?, ?, ?)''',
                  (row['CMIS ID'], row['Student Name'], row['CMIS PH No(10 Number)'],
                   row['Center Name'], row['Name Of Uploder'], row['Verification Type'], row['Mode Of Verification']))
    conn.commit()
    conn.close()
    st.success('Student data updated successfully!')

def insert_booking(date, time_range, manager, spoc, booked_by):
    """Insert slot booking into the slot booking database."""
    if not booked_by:
        st.error('Booking failed. Please provide your name.')
        return

    conn = sqlite3.connect(SLOT_BOOKING_DB)
    c = conn.cursor()
    
    c.execute('SELECT * FROM appointment_bookings WHERE date = ? AND spoc = ?', (date, spoc))
    existing_booking = c.fetchone()
    
    if existing_booking:
        conn.close()
        st.error('Slot already booked for this SPOC on the selected date.')
        return

    c.execute('''INSERT INTO appointment_bookings (date, time_range, manager, spoc, booked_by)
                 VALUES (?, ?, ?, ?, ?)''', (date, time_range, manager, spoc, booked_by))
    conn.commit()
    conn.close()
    st.success('Slot booked successfully!')

def download_student_data():
    """Download student data as CSV."""
    conn = sqlite3.connect(STUDENT_DB)
    df = pd.read_sql_query("SELECT * FROM studentcap", conn)
    conn.close()

    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="studentcap.csv">Download Student Data CSV</a>'
    st.markdown(href, unsafe_allow_html=True)

def fetch_data_from_databases():
    """Fetch data from both databases and return as DataFrames."""
    try:
        # Fetch student data
        conn_student = sqlite3.connect(STUDENT_DB)
        student_df = pd.read_sql_query("SELECT * FROM studentcap", conn_student)
        conn_student.close()
    except Exception as e:
        st.error(f"Error fetching student data: {e}")
        student_df = pd.DataFrame()

    try:
        # Fetch slot booking data
        conn_slot = sqlite3.connect(SLOT_BOOKING_DB)
        slot_df = pd.read_sql_query("SELECT * FROM appointment_bookings", conn_slot)
        conn_slot.close()
    except Exception as e:
        st.error(f"Error fetching slot booking data: {e}")
        slot_df = pd.DataFrame()

    return student_df, slot_df

def create_combined_excel(student_df, slot_df):
    """Create a combined Excel file and return the BytesIO object."""
    output = BytesIO()

    # Write to Excel
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not student_df.empty:
            student_df.to_excel(writer, index=False, sheet_name='Student Data')
        else:
            st.warning("No data available for 'Student Data' sheet.")

        if not slot_df.empty:
            slot_df.to_excel(writer, index=False, sheet_name='Slot Bookings')
        else:
            st.warning("No data available for 'Slot Bookings' sheet.")

    return output

def download_link(excel_file):
    """Generate download link for the Excel file."""
    processed_data = excel_file.getvalue()
    b64 = base64.b64encode(processed_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="combined_data.xlsx">Download Combined Excel File</a>'
    st.markdown(href, unsafe_allow_html=True)

# Main App UI
def main():
    st.title('Slot Booking & Student Data Management')

    # Create databases if not already present
    create_databases()

    # Upload and Update Student Data
    st.subheader('Upload Student Data for SPOC Calling')
    student_file = st.file_uploader('Upload Excel', type=['xlsx', 'xls'])
    if student_file:
        if st.button('Update Student Data'):
            update_student_database(student_file)

    # Download Student Data
    if st.button('Download Student Data'):
        download_student_data()

    # Slot Booking Section
    st.subheader('Slot Booking Platform')

    # Manager, SPOC, Date, Time, and Booking Fields
    manager = st.text_input('Manager Name')
    spoc = st.text_input('SPOC Name')
    date = st.date_input('Select Date')
    time_ranges = ['10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM', '2:00 PM - 3:00 PM']
    time_range = st.selectbox('Select Time Range', time_ranges)
    booked_by = st.text_input('Slot Booked By')

    # Book Slot Button
    if st.button('Book Slot'):
        insert_booking(str(date), time_range, manager, spoc, booked_by)

    # Fetch and display data on button click
    if st.button('Generate and Download Combined Excel'):
        student_df, slot_df = fetch_data_from_databases()

        if student_df.empty and slot_df.empty:
            st.warning("No data available in both databases.")
        else:
            excel_file = create_combined_excel(student_df, slot_df)
            download_link(excel_file)

if __name__ == '__main__':
    main()
