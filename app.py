import streamlit as st
import pandas as pd
import sqlite3
import calendar
from datetime import datetime, timedelta
import base64
import os
import shutil
from io import BytesIO

# Function to backup SQLite databases and save CSV backups
def backup_databases():
    # Ensure backup folder exists
    backup_folder = 'database_backups'
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
    
    # Backup slot_booking_new.db
    source_file1 = 'slot_booking_new.db'
    backup_file1 = f"{backup_folder}/slot_booking_new_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    
    if os.path.exists(source_file1):
        shutil.copy(source_file1, backup_file1)
    
    # Backup duplicate.db
    source_file2 = 'duplicate.db'
    backup_file2 = f"{backup_folder}/duplicate_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    
    if os.path.exists(source_file2):
        shutil.copy(source_file2, backup_file2)

    # Backup data to CSV for additional safety
    csv_backup_folder = f"{backup_folder}/csv_backups"
    if not os.path.exists(csv_backup_folder):
        os.makedirs(csv_backup_folder)

    # Backup studentcap table from duplicate.db to CSV
    conn_csv = sqlite3.connect('duplicate.db')
    df = pd.read_sql_query("SELECT * FROM studentcap", conn_csv)
    conn_csv.close()

    if not df.empty:
        csv_backup_file = f"{csv_backup_folder}/studentcap_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(csv_backup_file, index=False)

    st.success('Databases and CSV backups created successfully!')

# Function to delete old backups older than 45 days
def delete_old_backups():
    backup_folder = 'database_backups'
    if os.path.exists(backup_folder):
        for filename in os.listdir(backup_folder):
            file_path = os.path.join(backup_folder, filename)
            if os.path.isfile(file_path):
                creation_time = datetime.fromtimestamp(os.path.getctime(file_path))
                if datetime.now() - creation_time > timedelta(days=45):
                    os.remove(file_path)

# Function to load data from Excel into a DataFrame with @st.cache_data
@st.cache_data(hash_funcs={pd.DataFrame: lambda _: None})
def load_data(file):
    df = pd.read_excel(file)
    df.rename(columns={'Actual_Manager_Column_Name': 'Manager Name', 'Actual_SPOC_Column_Name': 'SPOC Name'}, inplace=True)
    return df

# Function to create SQLite database table for appointments
def create_table():
    conn = sqlite3.connect('slot_booking_new.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS appointment_bookings
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                 date TEXT,
                 time_range TEXT,
                 manager TEXT,
                 spoc TEXT,
                 booked_by TEXT)''')
    conn.commit()
    conn.close()

# Function to insert booking into SQLite database
def insert_booking(date, time_range, manager, spoc, booked_by):
    st.write(f'Attempting to book slot for: Date: {date}, Time Range: {time_range}, Manager: {manager}, SPOC: {spoc}, Booked By: {booked_by}')
    
    if not booked_by:
        st.error('Slot booking failed. You must provide your name in the "Slot Booked By" field.')
        return

    selected_date = datetime.strptime(date, '%Y-%m-%d')
    current_date = datetime.now()

    if selected_date < current_date:
        st.error('Slot booking failed. You cannot book slots for past dates.')
        return

    if selected_date.weekday() == 6:
        st.error('If Error Message Reflects Or To Book Slot On Holidays & Other Than Official Hours Please Contact To Pritam Basu & Kousik Dey.')
        return

    conn = sqlite3.connect('slot_booking_new.db')
    c = conn.cursor()

    c.execute('''SELECT * FROM appointment_bookings 
                 WHERE date = ? AND spoc = ?''', (date, spoc))
    existing_booking = c.fetchone()

    if existing_booking:
        conn.close()
        st.error('Slot booking failed. This SPOC is already booked for the selected date.')
        return

    c.execute('''INSERT INTO appointment_bookings (date, time_range, manager, spoc, booked_by)
                 VALUES (?, ?, ?, ?, ?)''', (date, time_range, manager, spoc, booked_by))
    conn.commit()
    conn.close()
    st.success('Slot booked successfully!')

# Function to update another database from uploaded Excel file
def update_another_database(file):
    df = pd.read_excel(file)

    conn = sqlite3.connect('duplicate.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS studentcap
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                 cmis_id TEXT,
                 student_name TEXT,
                 cmis_ph_no TEXT,
                 center_name TEXT,
                 uploader_name TEXT,
                 verification_type TEXT,
                 mode_of_verification TEXT)''')
    conn.commit()

    for index, row in df.iterrows():
        c.execute('''INSERT INTO studentcap (cmis_id, student_name, cmis_ph_no, center_name, uploader_name, verification_type, mode_of_verification)
                     VALUES (?, ?, ?, ?, ?, ?, ?)''', (row['CMIS ID'], row['Student Name'], row['CMIS PH No(10 Number)'],
                                                row['Center Name'], row['Name Of Uploder'], row['Verification Type'], row['Mode Of Verification']))
    conn.commit()
    conn.close()

    st.success('Data updated successfully!')

# Function to download data from duplicate.db
def download_another_database_data():
    conn = sqlite3.connect('duplicate.db')
    df = pd.read_sql_query("SELECT * FROM studentcap", conn)
    conn.close()
    
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="studentcap.csv">Download CSV</a>'
    st.markdown(href, unsafe_allow_html=True)

# Function to generate HTML for calendar view with bookings highlighted
def generate_calendar(bookings):
    cal = calendar.Calendar()
    current_year = datetime.now().year
    current_month = datetime.now().month
    weekday_names = list(calendar.day_abbr)

    days_html = ''
    for day in cal.itermonthdays(current_year, current_month):
        if day == 0:
            days_html += '<div class="day"></div>'
        else:
            date = pd.Timestamp(year=current_year, month=current_month, day=day)
            bookings_on_day = bookings[(bookings['date'].dt.year == current_year) &
                                       (bookings['date'].dt.month == current_month) &
                                       (bookings['date'].dt.day == day)]

            if date.weekday() == 6:
                day_style = 'background-color: red;'
            elif not bookings_on_day.empty:
                day_style = 'background-color: #b3e6b3;'
            else:
                day_style = ''

            days_html += f'<div class="day" style="{day_style}"><span class="day-number">{day}</span><br>{weekday_names[date.weekday()]}</div>'

    calendar_html = f"""
    <style>
        .calendar {{
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 10px;
            margin-top: 20px;
        }}
        .day {{
            padding: 10px;
            border: 1px solid #ccc;
            text-align: center;
        }}
        .booking {{
            background-color: #b3e6b3;
            padding: 5px;
            font-size: 0.8em;
        }}
        .day-number {{
            font-size: 1.2em;
            font-weight: bold;
        }}
    </style>
    <div class="calendar">
        {days_html}
    </div>
    """

    return calendar_html

# Function to bulk delete data from studentcap table in duplicate.db by cmis_id
def bulk_delete_studentcap(cmis_ids):
    conn = sqlite3.connect('duplicate.db')
    c = conn.cursor()
    
    for cmis_id in cmis_ids:
        c.execute("DELETE FROM studentcap WHERE cmis_id = ?", (cmis_id,))
    
    conn.commit()
    conn.close()
    st.success("Selected records deleted successfully.")

# Function to download sample Excel file
def download_sample_excel():
    # Sample data for the Excel file
    sample_data = {
        'CMIS ID': ['123', '456', '789'],
        'Student Name': ['John Doe', 'Jane Smith', 'Jim Beam'],
        'CMIS PH No(10 Number)': ['1234567890', '0987654321', '1122334455'],
        'Center Name': ['Center 1', 'Center 2', 'Center 3'],
        'Name Of Uploder': ['Uploader 1', 'Uploader 2', 'Uploader 3'],
        'Verification Type': ['Enrollment', 'Enrollment', 'Enrollment', ],
        'Mode Of Verification': ['G-meet', 'Call', 'MSG']
    }

    # Creating a DataFrame from the sample data
    df = pd.DataFrame(sample_data)

    # Creating a BytesIO object to store the Excel file
    excel_file = BytesIO()

    # Writing the DataFrame to the BytesIO object as an Excel file
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Resetting the position of the BytesIO object to the beginning
    excel_file.seek(0)

    # Creating a base64 encoded link to download the Excel file
    b64 = base64.b64encode(excel_file.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="sample_excel.xlsx">Download sample Excel file</a>'

    return href

# Main Streamlit application
def main():
    st.title('Database Management and Operations')

    st.sidebar.title('Options')
    selected_option = st.sidebar.selectbox('Select an option', ('Home', 'Book Slot', 'Manage Data', 'Backup & Delete'))

    if selected_option == 'Home':
        st.header('Welcome to Database Management Portal')
        st.write('Please select an option from the sidebar.')

    elif selected_option == 'Book Slot':
        st.header('Book a Slot')

        create_table()

        # Date selection
        st.subheader('Select Date')
        date = st.date_input('Date')

        # Time range selection
        st.subheader('Select Time Range')
        time_range_options = ['10:00 AM - 12:00 PM', '2:00 PM - 4:00 PM']
        time_range = st.selectbox('Time Range', time_range_options)

        # Manager and SPOC selection
        st.subheader('Enter Manager and SPOC Details')
        manager = st.text_input('Manager Name')
        spoc = st.text_input('SPOC Name')

        # Booked by
        st.subheader('Enter Your Name')
        booked_by = st.text_input('Slot Booked By')

        # Submit button to book slot
        if st.button('Book Slot'):
            insert_booking(date.strftime('%Y-%m-%d'), time_range, manager, spoc, booked_by)

    elif selected_option == 'Manage Data':
        st.header('Manage Data')

        # Upload Excel file to update another database
        st.subheader('Update Another Database from Excel File')
        file = st.file_uploader('Upload Excel file', type=['xlsx', 'xls'])
        if file is not None:
            if st.button('Update Database'):
                update_another_database(file)

        # Download data from another database
        st.subheader('Download Data from Another Database')
        if st.button('Download Data'):
            download_another_database_data()

        # Bulk delete from another database
        st.subheader('Bulk Delete from Another Database')
        cmis_ids = st.text_area('Enter CMIS IDs to delete (comma-separated)', height=100)
        cmis_ids = [x.strip() for x in cmis_ids.split(',')]
        if st.button('Delete Records'):
            bulk_delete_studentcap(cmis_ids)

    elif selected_option == 'Backup & Delete':
        st.header('Backup and Delete')

        # Backup databases
        if st.button('Backup Databases'):
            backup_databases()

        # Delete old backups
        if st.button('Delete Old Backups'):
            delete_old_backups()

    # Footer with sample Excel download link
    st.sidebar.markdown('---')
    st.sidebar.subheader('Download Sample Excel File')
    st.sidebar.markdown(download_sample_excel(), unsafe_allow_html=True)

    # Calendar view of bookings
    st.sidebar.subheader('Calendar View')
    conn = sqlite3.connect('slot_booking_new.db')
    bookings = pd.read_sql_query("SELECT * FROM appointment_bookings", conn)
    conn.close()
    st.sidebar.markdown(generate_calendar(bookings), unsafe_allow_html=True)

if __name__ == '__main__':
    main()
