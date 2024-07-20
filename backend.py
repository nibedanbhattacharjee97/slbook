import streamlit as st
import sqlite3
from datetime import datetime

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

# Main function for the Streamlit app
def main():
    st.title('Backend Slot Booking Platform')

    # Ensure table exists in SQLite database
    create_table()

    # Manager input
    manager = st.text_input('Enter Manager Name')

    # SPOC input
    spoc = st.text_input('Enter SPOC Name')

    # Date selection
    date = st.date_input('Select Date', value=datetime.today())

    # Time range selection
    time_ranges = ['10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM',
                   '12:00 PM - 1:00 PM', '2:00 PM - 3:00 PM', '3:00 PM - 4:00 PM']
    time_range = st.selectbox('Select Time', time_ranges)

    # Booked by (user input)
    booked_by = st.text_input('Slot Booked By')

    # Book button
    if st.button('Book Slot'):
        insert_booking(str(date), time_range, manager, spoc, booked_by)

# Run the app
if __name__ == '__main__':
    main()
