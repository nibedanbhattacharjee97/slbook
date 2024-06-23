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
    if not booked_by:
        st.error('Slot booking failed. You must provide your name in the "Your Name" field.')
        return

    selected_date = datetime.strptime(date, '%Y-%m-%d')
    current_date = datetime.now()

    if selected_date < current_date:
        st.error('Slot booking failed. You cannot book slots for past dates.')
        return

    if selected_date.weekday() == 6:
        st.error('Slot booking failed. Booking slots on Sundays is not allowed, For Special Permissions Call Pritam Basu Or Kousik Dey.')
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

    conn = sqlite3.connect('another_database.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS student_data
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                 cmis_id TEXT,
                 student_name TEXT,
                 cmis_ph_no TEXT,
                 center_name TEXT,
                 uploader_name TEXT)''')
    conn.commit()

    for index, row in df.iterrows():
        c.execute('''INSERT INTO student_data (cmis_id, student_name, cmis_ph_no, center_name, uploader_name)
                     VALUES (?, ?, ?, ?, ?)''', (row['CMIS ID'], row['Student Name'], row['CMIS PH No(10 Number)'],
                                                row['Center Name'], row['Name Of Uploder']))
    conn.commit()
    conn.close()

    st.success('Data updated successfully!')

# Function to download data from another_database.db
def download_another_database_data():
    conn = sqlite3.connect('another_database.db')
    df = pd.read_sql_query("SELECT * FROM student_data", conn)
    conn.close()
    
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="student_data.csv">Download CSV</a>'
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

# Function to bulk delete data from student_data table in another_database.db by cmis_id
def bulk_delete_student_data(cmis_ids):
    conn = sqlite3.connect('another_database.db')
    c = conn.cursor()
    
    for cmis_id in cmis_ids:
        c.execute("DELETE FROM student_data WHERE cmis_id = ?", (cmis_id,))
    
    conn.commit()
    conn.close()
    st.success("Selected records deleted successfully.")

# Function to delete booking by ID
def delete_booking_by_id(booking_id, user_email):
    allowed_email = "nibedan.b@anudip.org"
    if user_email != allowed_email:
        st.error("You do not have permission to delete booking data.")
        return

    conn = sqlite3.connect('slot_booking_new.db')
    c = conn.cursor()
    c.execute("DELETE FROM appointment_bookings WHERE id = ?", (booking_id,))
    conn.commit()
    conn.close()
    st.success(f"Booking with ID {booking_id} deleted successfully.")

# Main function for the Streamlit app
def main():
    st.title('Slot Booking Platform')

    # Ensure table exists in SQLite database
    create_table()

    # Load data using st.cache_data
    data = load_data('managers_spocs.xlsx')

    # Date selection
    selected_date = st.date_input('Select Date')

    # Time range selection
    time_ranges = ['10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM',
                   '12:00 PM - 1:00 PM', '2:00 PM - 3:00 PM', '3:00 PM - 4:00 PM']
    selected_time_range = st.selectbox('Select Time', time_ranges)

    # Manager selection
    selected_manager = st.selectbox('Select Manager', data['Manager Name'].unique())

    # SPOC selection based on selected manager
    spocs_for_manager = data[data['Manager Name'] == selected_manager]['SPOC Name'].tolist()
    selected_spoc = st.selectbox('Select SPOC', spocs_for_manager)

    # Booked by (user input)
    booked_by = st.text_input('Slot Booker Name')

    # Book button
    if st.button('Book Slot'):
        insert_booking(str(selected_date), selected_time_range, selected_manager, selected_spoc, booked_by)

    # Upload Excel file and update another database
    st.subheader('Upload Excel for another database update')
    file = st.file_uploader('Upload Excel', type=['xlsx', 'xls'])
    if file is not None:
        if st.button('Update Another Database'):
            update_another_database(file)

    # Download data button
    if st.button('Download Data from Another Database'):
        download_another_database_data()

    # Fetch all bookings
    conn = sqlite3.connect('slot_booking_new.db')
    bookings = pd.read_sql_query("SELECT * FROM appointment_bookings", conn)
    conn.close()

    if 'date' in bookings.columns:
        bookings['date'] = pd.to_datetime(bookings['date'])

    # Show calendar after booking attempt
    st.subheader('Calendar View (Current Month Status)')
    st.markdown(generate_calendar(bookings), unsafe_allow_html=True)

    # Display today's bookings
    st.header("Today's Bookings")

    current_date = datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect('slot_booking_new.db')
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
        b64 = base64.b64encode(csv.encode()).decode()
        href = f'<a href="data:file/csv;base64,{b64}" download="monthly_bookings.csv">Download CSV</a>'
        st.markdown(href, unsafe_allow_html=True)

    # Bulk delete student data
    st.header('Bulk Delete Student Data')
    file = st.file_uploader('Upload CSV with CMIS IDs to delete', type=['csv'])
    if file is not None:
        cmis_ids_df = pd.read_csv(file)
        cmis_ids = cmis_ids_df['cmis_id'].tolist()
        if st.button('Delete Records'):
            bulk_delete_student_data(cmis_ids)

    # Delete booking by ID
    st.header("Delete Booking by ID")
    booking_id = st.number_input("Enter Booking ID to delete", min_value=1)
    user_email = st.text_input("Enter your email")
    if st.button("Delete Booking"):
        delete_booking_by_id(booking_id, user_email)

# Run the app
if __name__ == '__main__':
    main()
