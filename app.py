import streamlit as st
import pandas as pd
import sqlite3
import calendar
from datetime import datetime
import base64
from io import BytesIO

#st.image("dddd.jpg", use_column_width=True)

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

    # List of holiday dates (for example, Durga Puja)
    holidays = ['2024-31-10','2024-09-11','2024-09-16']  # Add dates in 'YYYY-MM-DD' format

    # Check if the selected date falls on the holiday
    if selected_date.strftime('%Y-%m-%d') in holidays:
        st.error('Booking Closed')
        return

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

def update_another_database(file):
    df = pd.read_excel(file)

    conn = sqlite3.connect('Plana.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS plana
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                 cmis_id TEXT,
                 student_name TEXT,
                 cmis_ph_no TEXT,
                 center_name TEXT,
                 uploader_name TEXT,
                 verification_type TEXT,
                 mode_of_verification TEXT,
                 verification_date TEXT)''')
    conn.commit()

    for index, row in df.iterrows():
        c.execute('''INSERT INTO plana (cmis_id, student_name, cmis_ph_no, center_name, uploader_name, verification_type, mode_of_verification, verification_date)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?)''', (row['CMIS ID'], row['Student Name'], row['CMIS PH No(10 Number)'],
                                                          row['Center Name'], row['Name Of Uploder'], row['Verification Type'], row['Mode Of Verification'], row['Verification Date']))
    conn.commit()
    conn.close()

    st.success('Data updated successfully!')

# Function to download data from Plana.db
def download_another_database_data():
    conn = sqlite3.connect('Plana.db')
    df = pd.read_sql_query("SELECT * FROM plana", conn)
    conn.close()
    
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="plana.csv">Download CSV</a>'
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

# Function to bulk delete data from plana table in Plana.db by cmis_id
def bulk_delete_plana(cmis_ids):
    conn = sqlite3.connect('Plana.db')
    c = conn.cursor()
    
    for cmis_id in cmis_ids:
        c.execute("DELETE FROM plana WHERE cmis_id = ?", (cmis_id,))
    
    conn.commit()
    conn.close()
    st.success("Selected records deleted successfully.")

def download_sample_excel():
    sample_data = {
        'CMIS ID': ['123', '456', '789'],
        'Student Name': ['John Doe', 'Jane Smith', 'Jim Beam'],
        'CMIS PH No(10 Number)': ['1234567890', '0987654321', '1122334455'],
        'Center Name': ['Center 1', 'Center 2', 'Center 3'],
        'Name Of Uploder': ['Uploader 1', 'Uploader 2', 'Uploader 3'],
        'Verification Type': ['Placement', 'Placement', 'Enrollment', ],
        'Mode Of Verification': ['G-meet', 'Call', 'Call'],
        'Verification Date': ['28-02-2025', '28-02-2025', '28-02-2025']
    }

    # Creating a DataFrame from the sample data
    df = pd.DataFrame(sample_data)

    # Creating a BytesIO object to store the Excel file
    output = BytesIO()

    # Using ExcelWriter with xlsxwriter engine to write the DataFrame to Excel format in the BytesIO object
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Get the xlsxwriter workbook and worksheet objects
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # Adjust column width
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)

    # Resetting the buffer position to the start of the buffer
    output.seek(0)

    # Encoding the Excel file in base64
    excel_file = output.read()
    b64 = base64.b64encode(excel_file).decode()

    # Creating a download link for the Excel file
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="Sample_Excel.xlsx">Download Sample Excel</a>'

    # Displaying the download link in Streamlit
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
            insert_booking(str(selected_date), selected_time_range, selected_manager, selected_spoc, booked_by)

    # Download sample Excel file
    st.subheader('Download The Format To Update Student Data For SPOC Calling')
    if st.button('Download Sample'):
        download_sample_excel()

    # Download data button
    if st.button('Download Data For M&E Purpose'):
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

## Show Excel

    st.subheader("SPOC List For Upcoming Dates")
    try:
        spoc_list_df = pd.read_excel("SPOC_List.xlsx")
        st.dataframe(spoc_list_df)
    except FileNotFoundError:
        st.error("SPOC_List.xlsx not found. Please ensure the file is in the correct directory.")
    except Exception as e:
        st.error(f"An error occurred while loading SPOC_List.xlsx: {e}")



# Run the app
if __name__ == '__main__':
    main()
