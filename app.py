import streamlit as st
import pandas as pd
import sqlite3
import calendar
from datetime import datetime
import base64
from io import BytesIO

#st.text("The Slot Booking Platform is currently under development,so dont Book Slot For Now.")

@st.cache_data(hash_funcs={pd.DataFrame: lambda _: None})
def load_data(file):
    df = pd.read_excel(file)
    df.rename(columns={'Actual_Manager_Column_Name': 'Manager Name', 'Actual_SPOC_Column_Name': 'SPOC Name'}, inplace=True)
    return df

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

def insert_booking(date, time_range, manager, spoc, booked_by):
    st.write(f'Attempting to book slot for: Date: {date}, Time Range: {time_range}, Manager: {manager}, SPOC: {spoc}, Booked By: {booked_by}')
    if not booked_by:
        st.error('Slot booking failed. You must provide your name in the "Slot Booked By" field.')
        return

    selected_date = datetime.strptime(date, '%Y-%m-%d')
    current_date = datetime.now()

    holidays = ['2024-31-10', '2024-09-11', '2024-09-16']
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

# Update plana.db only with CMIS_IDs present in ids.xlsx
def update_another_database(file):
    df = pd.read_excel(file)
    ids_df = pd.read_excel('ids.xlsx')
    valid_ids = ids_df['CMIS_ID'].astype(str).unique()

    df['CMIS ID'] = df['CMIS ID'].astype(str)
    filtered_df = df[df['CMIS ID'].isin(valid_ids)]

    if filtered_df.empty:
        st.error("")
        return

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

    for _, row in filtered_df.iterrows():
        c.execute('''INSERT INTO plana (cmis_id, student_name, cmis_ph_no, center_name, uploader_name, verification_type, mode_of_verification, verification_date)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?)''', (row['CMIS ID'], row['Student Name'], row['CMIS PH No(10 Number)'],
                                                          row['Center Name'], row['Name Of Uploder'], row['Verification Type'], row['Mode Of Verification'], row['Verification Date']))
    conn.commit()
    conn.close()
    st.success(f"{len(filtered_df)} valid records inserted successfully.")

def download_another_database_data():
    conn = sqlite3.connect('Plana.db')
    df = pd.read_sql_query("SELECT * FROM plana", conn)
    conn.close()

    ids_df = pd.read_excel('ids.xlsx')
    valid_ids = ids_df['CMIS_ID'].astype(str).unique()
    df['cmis_id'] = df['cmis_id'].astype(str)

    filtered_df = df[df['cmis_id'].isin(valid_ids)]

    if filtered_df.empty:
        st.error("No valid data found for M&E verification.")
        return

    csv = filtered_df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="plana_filtered.csv">Download CSV</a>'
    st.markdown(href, unsafe_allow_html=True)

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
            day_style = 'background-color: red;' if date.weekday() == 6 else ('background-color: #b3e6b3;' if not bookings_on_day.empty else '')
            days_html += f'<div class="day" style="{day_style}"><span class="day-number">{day}</span><br>{weekday_names[date.weekday()]}</div>'

    return f"""
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
        .day-number {{
            font-size: 1.2em;
            font-weight: bold;
        }}
    </style>
    <div class="calendar">
        {days_html}
    </div>
    """

def download_sample_excel():
    sample_data = {
        'CMIS ID': ['123', '456', '789'],
        'Student Name': ['John Doe', 'Jane Smith', 'Jim Beam'],
        'CMIS PH No(10 Number)': ['1234567890', '0987654321', '1122334455'],
        'Center Name': ['Center 1', 'Center 2', 'Center 3'],
        'Name Of Uploder': ['Uploader 1', 'Uploader 2', 'Uploader 3'],
        'Verification Type': ['Placement', 'Placement', 'Enrollment'],
        'Mode Of Verification': ['G-meet', 'Call', 'Call'],
        'Verification Date': ['28-02-2025', '28-02-2025', '28-02-2025']
    }

    df = pd.DataFrame(sample_data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="Sample_Excel.xlsx">Download Sample Excel</a>'
    st.markdown(href, unsafe_allow_html=True)

def main():
    st.title('Slot Booking Platform')
    create_table()
    data = load_data('managers_spocs.xlsx')

    selected_manager = st.selectbox('Select Manager', data['Manager Name'].unique())
    spocs_for_manager = data[data['Manager Name'] == selected_manager]['SPOC Name'].tolist()
    selected_spoc = st.selectbox('Select SPOC', spocs_for_manager)

    selected_date = st.date_input('Select Date')
    time_ranges = ['10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM', '12:00 PM - 1:00 PM', '2:00 PM - 3:00 PM', '3:00 PM - 4:00 PM']
    selected_time_range = st.selectbox('Select Time', time_ranges)
    booked_by = st.text_input('Slot Booked By')

    st.subheader('Upload Student Data For SPOC Calling')

    st.subheader('Please ensure that only student data marked as Not Joined or Not Contacted from the M&E database is included.')

    file = st.file_uploader('Upload Excel', type=['xlsx', 'xls'])
    data_uploaded = st.session_state.get('data_uploaded', False)
    if file is not None:
        if st.button('Update Data'):
            update_another_database(file)
            st.session_state['data_uploaded'] = True
            data_uploaded = True

    if not data_uploaded:
        st.warning('Please upload student data before booking a slot.')
    else:
        if st.button('Book Slot'):
            insert_booking(str(selected_date), selected_time_range, selected_manager, selected_spoc, booked_by)

    st.subheader('Download The Format To Update Student Data For SPOC Calling')
    if st.button('Download Sample'):
        download_sample_excel()

    if st.button('Download Data For M&E Purpose'):
        download_another_database_data()

    conn = sqlite3.connect('slot_booking_new.db')
    bookings = pd.read_sql_query("SELECT * FROM appointment_bookings", conn)
    conn.close()
    if 'date' in bookings.columns:
        bookings['date'] = pd.to_datetime(bookings['date'])

    st.subheader('Calendar View (Current Month Status)')
    st.markdown(generate_calendar(bookings), unsafe_allow_html=True)

    st.header("Today's Bookings")
    current_date = datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect('slot_booking_new.db')
    c = conn.cursor()
    c.execute("SELECT date, time_range, manager, spoc FROM appointment_bookings WHERE date = ?", (current_date,))
    today_booking_details = c.fetchall()
    conn.close()

    if today_booking_details:
        st.write(f"Bookings for today ({current_date}):")
        for detail in today_booking_details:
            st.write(f"- Time Slot: {detail[1]}, Manager: {detail[2]}, SPOC: {detail[3]}")
    else:
        st.write("No bookings for today.")

    if st.button('Download Monthly Data'):
        csv = bookings.to_csv(index=False)
        b64 = base64.b64encode(csv.encode()).decode()
        href = f'<a href="data:file/csv;base64,{b64}" download="monthly_bookings.csv">Download CSV</a>'
        st.markdown(href, unsafe_allow_html=True)

if __name__ == '__main__':
    main()
