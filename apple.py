import streamlit as st
import pandas as pd
import sqlite3
import calendar
from datetime import datetime, timedelta
import base64
from io import BytesIO
import xlsxwriter

# --- Function Definitions ---

def load_data(file):
    df = pd.read_excel(file)
    df.rename(columns={'Actual_Manager_Column_Name': 'Manager Name', 
                       'Actual_SPOC_Column_Name': 'SPOC Name'}, inplace=True)  # Replace with your actual column names
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
    # ... (validation code remains the same)

    conn = sqlite3.connect('slot_booking_new.db')
    c = conn.cursor()

    # ... (existing booking check remains the same)

    c.execute('''INSERT INTO appointment_bookings (date, time_range, manager, spoc, booked_by)
                 VALUES (?, ?, ?, ?, ?)''', (date, time_range, manager, spoc, booked_by))
    conn.commit()
    conn.close()
    st.success('Slot booked successfully!')


def update_another_database(file):
    df = pd.read_excel(file)

    conn = sqlite3.connect('duplicate.db')
    c = conn.cursor()

    # ***KEY CHANGE: DELETE existing data before inserting new data***
    c.execute("DELETE FROM studentcap")  # Clear the table
    conn.commit()

    # Now insert the new data
    for index, row in df.iterrows():
        c.execute('''INSERT INTO studentcap (cmis_id, student_name, cmis_ph_no, center_name, uploader_name, verification_type, mode_of_verification)
                     VALUES (?, ?, ?, ?, ?, ?, ?)''', (row['CMIS ID'], row['Student Name'], row['CMIS PH No(10 Number)'],
                                                          row['Center Name'], row['Name Of Uploder'], row['Verification Type'], 
                                                          row['Mode Of Verification']))
    conn.commit()
    conn.close()

    st.success('Data updated successfully!')

def download_another_database_data():
    conn = sqlite3.connect('duplicate.db')
    df = pd.read_sql_query("SELECT * FROM studentcap", conn)
    conn.close()

    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="studentcap.csv">Download CSV</a>'
    st.markdown(href, unsafe_allow_html=True)


def download_combined_data():
    try:
        conn1 = sqlite3.connect('slot_booking_new.db')
        bookings_df = pd.read_sql_query("SELECT * FROM appointment_bookings", conn1)
        conn1.close()

        conn2 = sqlite3.connect('duplicate.db')
        studentcap_df = pd.read_sql_query("SELECT * FROM studentcap", conn2)
        conn2.close()

        # ***The Correct Merge Logic (Improved)***
        bookings_df['dummy'] = 1
        studentcap_df['dummy'] = 1

        merged_df = pd.merge(bookings_df, studentcap_df, on='dummy', how='outer')

        # Dynamic date column handling (Improved)
        date_column = next((col for col in merged_df.columns if 'date' in col.lower()), None)
        spoc_column_studentcap = next((col for col in studentcap_df.columns if 'spoc' in col.lower() or 'student_name' in col.lower()), None) # added to handle different column names for spoc

        combined_df = merged_df[
            (merged_df['date'] == merged_df[date_column]) &
            (merged_df['spoc'] == merged_df[spoc_column_studentcap])  # Use the dynamic date and spoc column
        ]

        combined_df = combined_df.drop('dummy', axis=1)

        # Convert date columns to string before exporting to Excel
        for col in combined_df.columns:
            if 'date' in col.lower():
                combined_df[col] = combined_df[col].astype(str)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Combined Data')

            workbook = writer.book
            worksheet = writer.sheets['Combined Data']

            # Adjust column width
            for i, col in enumerate(combined_df.columns):
                column_len = max(combined_df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_len)

        output.seek(0)
        excel_file = output.read()
        b64 = base64.b64encode(excel_file).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="combined_data.xlsx">Download Combined Data</a>'
        st.markdown(href, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"An error occurred during data download: {e}")


# ... (generate_calendar, bulk_delete_studentcap, download_sample_excel remain the same)



def main():
    st.title('Slot Booking Platform')

    create_table()

    data = load_data('managers_spocs.xlsx')

    selected_manager = st.selectbox('Select Manager', data['Manager Name'].unique())

    spocs_for_manager = data[data['Manager Name'] == selected_manager]['SPOC Name'].tolist()
    selected_spoc = st.selectbox('Select SPOC', spocs_for_manager)

    selected_date = st.date_input('Select Date')
    selected_date_str = str(selected_date)  # Convert to string for database

    time_ranges = ['10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM',
                   '12:00 PM - 1:00 PM', '2:00 PM - 3:00 PM', '3:00 PM - 4:00 PM']
    selected_time_range = st.selectbox('Select Time', time_ranges)

    booked_by = st.text_input('Slot Booked By')

    # Initialize data_uploaded in session state
    if 'data_uploaded' not in st.session_state:
        st.session_state.data_uploaded = False
    data_uploaded = st.session_state.data_uploaded

    file = st.file_uploader('Upload Excel', type=['xlsx', 'xls'])
    if file is not None:
        if st.button('Update Data'):
            update_another_database(file)
            st.session_state.data_uploaded = True
            data_uploaded = True


    if not data_uploaded:
        st.warning('Please upload student data before booking a slot.')
    else:
        if st.button('Book Slot'):
            insert_booking(selected_date_str, selected_time_range, selected_manager, selected_spoc, booked_by)

    # ... (rest of the calendar, today's bookings, monthly download, bulk delete remain the same)

    if st.button('Download Combined Data'):
        download_combined_data()


if __name__ == '__main__':
    main()