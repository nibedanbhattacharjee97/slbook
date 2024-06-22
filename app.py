import streamlit as st
import pandas as pd
import sqlite3
import calendar
from datetime import datetime
import base64

# Function to load data from Excel into a DataFrame with @st.cache_data
@st.cache_data(hash_funcs={pd.DataFrame: lambda _: None})
def load_data(file):
    df = pd.read_excel(file)
    # Assuming 'Manager Name' and 'SPOC Name' are the actual column names
    df.rename(columns={'Actual_Manager_Column_Name': 'Manager Name', 'Actual_SPOC_Column_Name': 'SPOC Name'}, inplace=True)
    # Perform any additional data processing within this function
    return df

# Function to create SQLite database table for appointments
def create_table():
    conn = sqlite3.connect('slot_booking_new.db')  # Change database name if needed
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
    # Check if the selected date is in the past
    selected_date = datetime.strptime(date, '%Y-%m-%d')
    current_date = datetime.now()

    if selected_date < current_date:
        st.error('Slot booking failed. You cannot book slots for past dates.')
        return

    conn = sqlite3.connect('slot_booking_new.db')  # Change database name if needed
    c = conn.cursor()

    # Check if there is already a booking for this SPOC on this date
    c.execute('''SELECT * FROM appointment_bookings 
                 WHERE date = ? AND spoc = ?''', (date, spoc))
    existing_booking = c.fetchone()

    if existing_booking:
        conn.close()
        st.error('Slot booking failed. This SPOC is already booked for the selected date.')
        return

    # Insert the new booking
    c.execute('''INSERT INTO appointment_bookings (date, time_range, manager, spoc, booked_by)
                 VALUES (?, ?, ?, ?, ?)''', (date, time_range, manager, spoc, booked_by))
    conn.commit()
    conn.close()
    st.success('Slot booked successfully!')

# Function to generate HTML for calendar view with bookings highlighted
def generate_calendar(bookings):
    cal = calendar.Calendar()
    current_year = datetime.now().year
    current_month = datetime.now().month

    # Get the weekday names
    weekday_names = list(calendar.day_abbr)  # Using abbreviated day names (Mon, Tue, Wed, etc.)

    days_html = ''
    for day in cal.itermonthdays(current_year, current_month):
        if day == 0:
            days_html += '<div class="day"></div>'
        else:
            date = pd.Timestamp(year=current_year, month=current_month, day=day)
            bookings_on_day = bookings[(bookings['date'].dt.year == current_year) &
                                       (bookings['date'].dt.month == current_month) &
                                       (bookings['date'].dt.day == day)]

            if not bookings_on_day.empty:
                days_html += '<div class="day booking"><span class="day-number">%d</span><br>%s</div>' % (
                    day, weekday_names[date.weekday()])
            else:
                days_html += '<div class="day"><span class="day-number">%d</span><br>%s</div>' % (
                    day, weekday_names[date.weekday()])

    calendar_html = """
    <style>
        .calendar {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 10px;
            margin-top: 20px;
        }
        .day {
            padding: 10px;
            border: 1px solid #ccc;
            text-align: center;
        }
        .booking {
            background-color: #b3e6b3;
            padding: 5px;
            font-size: 0.8em;
        }
        .day-number {
            font-size: 1.2em;
            font-weight: bold;
        }
    </style>
    <div class="calendar">
        %s
    </div>
    """ % days_html

    return calendar_html

# Main function for the Streamlit app
def main():
    st.title('Slot Booking Platform')

    # Ensure table exists in SQLite database
    create_table()

    # Load data using st.cache_data
    data = load_data('managers_spocs.xlsx')

    # Date selection
    selected_date = st.date_input('Select a date')

    # Time range selection
    time_ranges = ['9:00 AM - 10:00 AM', '10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM',
                   '12:00 PM - 1:00 PM', '2:00 PM - 3:00 PM', '3:00 PM - 4:00 PM']
    selected_time_range = st.selectbox('Select a time range', time_ranges)

    # Manager selection
    selected_manager = st.selectbox('Select a manager', data['Manager Name'].unique())

    # SPOC selection based on selected manager
    spocs_for_manager = data[data['Manager Name'] == selected_manager]['SPOC Name'].tolist()
    selected_spoc = st.selectbox('Select a SPOC', spocs_for_manager)

    # Booked by (user input)
    booked_by = st.text_input('Your Name')

    # Book button
    if st.button('Book Slot'):
        insert_booking(str(selected_date), selected_time_range, selected_manager, selected_spoc, booked_by)

    # Fetch all bookings
    conn = sqlite3.connect('slot_booking_new.db')  # Change database name if needed
    bookings = pd.read_sql_query("SELECT * FROM appointment_bookings", conn)
    conn.close()

    # Convert 'date' column to datetime if it's stored as text
    if 'date' in bookings.columns:
        bookings['date'] = pd.to_datetime(bookings['date'])

    # Show calendar after booking attempt
    st.subheader('Calendar View')
    st.markdown(generate_calendar(bookings), unsafe_allow_html=True)

    # Display today's bookings
    st.header("Today's Bookings")

    current_date = datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect('slot_booking_new.db')  # Change database name if needed
    c = conn.cursor()
    c.execute(
        "SELECT date, time_range, manager, spoc FROM appointment_bookings WHERE date = ?",
        (current_date,)
    )
    today_booking_details = c.fetchall()
    conn.close()

    if today_booking_details:
        st.write(f"Bookings for today ({current_date}):")
        for detail in today_booking_details:
            st.write(f"- Time Slot: {detail[1]}, Manager: {detail[2]}, SPOC: {detail[3]}")
    else:
        st.write("No bookings for today.")

    # Download button for monthly data
    if st.button('Download Monthly Data'):
        csv = bookings.to_csv(index=False)
        b64 = base64.b64encode(csv.encode()).decode()  # B64 encoding
        href = f'<a href="data:file/csv;base64,{b64}" download="monthly_bookings.csv">Download CSV</a>'
        st.markdown(href, unsafe_allow_html=True)

# Run the app
if __name__ == '__main__':
    main()
