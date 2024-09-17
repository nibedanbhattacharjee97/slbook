import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import base64

# Function to load data from Excel into a DataFrame with @st.cache_data
@st.cache_data(hash_funcs={pd.DataFrame: lambda _: None})
def load_data(file):
    df = pd.read_excel(file)
    df.rename(columns={'Actual_Manager_Column_Name': 'Manager Name', 'Actual_SPOC_Column_Name': 'SPOC Name'}, inplace=True)
    return df

# Function to create SQLite database table for appointments with the correct schema
def create_table():
    conn = sqlite3.connect('slot_booking_new.db')
    c = conn.cursor()

    # Create table for slot bookings
    c.execute('''CREATE TABLE IF NOT EXISTS appointment_bookings
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                 booking_id TEXT,
                 date TEXT,
                 time_range TEXT,
                 manager TEXT,
                 spoc TEXT,
                 booked_by TEXT)''')

    conn.commit()
    conn.close()

# Function to insert booking into SQLite database
def insert_booking(booking_id, date, time_range, manager, spoc, booked_by):
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

    c.execute('''INSERT INTO appointment_bookings (booking_id, date, time_range, manager, spoc, booked_by)
                 VALUES (?, ?, ?, ?, ?, ?)''', (booking_id, date, time_range, manager, spoc, booked_by))
    conn.commit()
    conn.close()
    st.success('Slot booked successfully!')

    # Update session state with booking ID
    st.session_state.booking_id = booking_id

# Function to update student data in the second database, linked by booking_id
def update_another_database(file):
    if 'booking_id' not in st.session_state:
        st.error('No booking ID found. Please book a slot first.')
        return

    booking_id = st.session_state.booking_id
    df = pd.read_excel(file)

    # Add the booking_id to the DataFrame
    df['booking_id'] = booking_id

    conn = sqlite3.connect('duplicate.db')
    c = conn.cursor()

    # Drop the existing table if it exists
    c.execute('DROP TABLE IF EXISTS studentcap')
    
    # Recreate the table with the correct schema
    c.execute('''CREATE TABLE IF NOT EXISTS studentcap
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                 booking_id TEXT,
                 cmis_id TEXT,
                 student_name TEXT,
                 cmis_ph_no TEXT,
                 center_name TEXT,
                 uploader_name TEXT,
                 verification_type TEXT,
                 mode_of_verification TEXT)''')
    conn.commit()

    for index, row in df.iterrows():
        c.execute('''INSERT INTO studentcap (booking_id, cmis_id, student_name, cmis_ph_no, center_name, uploader_name, verification_type, mode_of_verification)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?)''', (row['booking_id'], row['CMIS ID'], row['Student Name'], row['CMIS PH No(10 Number)'],
                                                row['Center Name'], row['Name Of Uploder'], row['Verification Type'], row['Mode Of Verification']))
    conn.commit()
    conn.close()

    st.success('Data updated successfully!')

# Function to merge data from both databases based on booking_id and download the result
def download_merged_data():
    conn1 = sqlite3.connect('slot_booking_new.db')
    conn2 = sqlite3.connect('duplicate.db')

    bookings = pd.read_sql_query("SELECT * FROM appointment_bookings", conn1)
    student_data = pd.read_sql_query("SELECT * FROM studentcap", conn2)

    conn1.close()
    conn2.close()

    if bookings.empty or student_data.empty:
        st.warning("No data available to merge.")
        return

    # Log the bookings and student data to help debug
    st.write("Bookings Data:")
    st.write(bookings)
    
    st.write("Student Data:")
    st.write(student_data)

    # Merge both DataFrames using booking_id
    merged_data = pd.merge(bookings, student_data, on='booking_id', how='inner')

    if merged_data.empty:
        st.warning("No matching data found between bookings and student data.")
        return

    # Generate CSV for download
    csv = merged_data.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="merged_data.csv">Download Merged Data</a>'
    st.markdown(href, unsafe_allow_html=True)

# Main function for the Streamlit app
def main():
    st.title('Slot Booking Platform')

    # Ensure table exists in SQLite database
    create_table()

    # Load data using st.cache_data
    data = load_data('managers_spocs.xlsx')

    # Manager selection
    selected_manager = st.selectbox('Select Manager', data['Manager Name'].unique())

    # SPOC selection based on selected manager
    spocs_for_manager = data[data['Manager Name'] == selected_manager]['SPOC Name'].tolist()
    selected_spoc = st.selectbox('Select SPOC', spocs_for_manager)

    # Date selection
    selected_date = st.date_input('Select Date')

    # Time range selection
    time_ranges = ['10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM',
                   '12:00 PM - 1:00 PM', '2:00 PM - 3:00 PM', '3:00 PM - 4:00 PM']
    selected_time_range = st.selectbox('Select Time', time_ranges)

    # Booked by (user input)
    booked_by = st.text_input('Slot Booked By')

    # Generate a unique booking ID
    if 'booking_id' not in st.session_state:
        st.session_state.booking_id = None

    # Upload Excel file and update another database
    st.subheader('Upload Student Data For SPOC Calling')
    file = st.file_uploader('Upload Excel', type=['xlsx', 'xls'])
    data_uploaded = st.session_state.get('data_uploaded', False)
    if file is not None:
        if st.button('Update Data'):
            update_another_database(file)
            st.session_state['data_uploaded'] = True
            data_uploaded = True

    # Only allow booking if data is uploaded
    if not data_uploaded:
        st.warning('Please upload student data before booking a slot.')
    else:
        # Book button
        if st.button('Book Slot'):
            insert_booking(str(datetime.now().timestamp()).replace('.', ''), str(selected_date), selected_time_range, selected_manager, selected_spoc, booked_by)

    # Download merged data
    if st.button('Download Merged Data'):
        download_merged_data()

# Run the app
if __name__ == '__main__':
    main()
