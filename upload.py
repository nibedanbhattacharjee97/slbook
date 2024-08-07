import sqlite3
import pandas as pd
import streamlit as st

# Function to upload and update slot_booking_new.db with CSV or Excel data
def upload_slot_booking(file):
    file_extension = file.name.split('.')[-1]
    
    # Read the file based on its extension
    if file_extension == 'csv':
        df = pd.read_csv(file)
    elif file_extension in ['xls', 'xlsx']:
        df = pd.read_excel(file)
    else:
        st.error("Unsupported file format")
        return
    
    # Connect to the SQLite database
    conn = sqlite3.connect('slot_booking_new.db')
    # Update the database with the file data
    df.to_sql('appointment_bookings', conn, if_exists='replace', index=False)
    conn.close()
    st.success('slot_booking_new.db updated successfully with uploaded data.')

# Function to upload and update duplicate.db with CSV or Excel data
def upload_duplicate(file):
    file_extension = file.name.split('.')[-1]
    
    # Read the file based on its extension
    if file_extension == 'csv':
        df = pd.read_csv(file)
    elif file_extension in ['xls', 'xlsx']:
        df = pd.read_excel(file)
    else:
        st.error("Unsupported file format")
        return
    
    # Connect to the SQLite database
    conn = sqlite3.connect('duplicate.db')
    # Update the database with the file data
    df.to_sql('studentcap', conn, if_exists='replace', index=False)
    conn.close()
    st.success('duplicate.db updated successfully with uploaded data.')

# Function to view data from slot_booking_new.db
def view_slot_booking_data():
    conn = sqlite3.connect('slot_booking_new.db')
    df = pd.read_sql_query("SELECT * FROM appointment_bookings", conn)
    conn.close()
    st.dataframe(df)

# Function to view data from duplicate.db
def view_duplicate_data():
    conn = sqlite3.connect('duplicate.db')
    df = pd.read_sql_query("SELECT * FROM studentcap", conn)
    conn.close()
    st.dataframe(df)

# Function to export data from slot_booking_new.db to CSV
def export_slot_booking_to_csv():
    conn = sqlite3.connect('slot_booking_new.db')
    df = pd.read_sql_query("SELECT * FROM appointment_bookings", conn)
    conn.close()
    return df.to_csv(index=False)

# Function to export data from duplicate.db to CSV
def export_duplicate_to_csv():
    conn = sqlite3.connect('duplicate.db')
    df = pd.read_sql_query("SELECT * FROM studentcap", conn)
    conn.close()
    return df.to_csv(index=False)

# Main function for the Streamlit app
def main():
    st.title('Database Upload and View')

    st.header('Upload Files to Update Databases')

    slot_booking_file = st.file_uploader('Upload slot_booking_new.csv or slot_booking_new.xlsx', type=['csv', 'xls', 'xlsx'])
    if slot_booking_file:
        upload_slot_booking(slot_booking_file)

    duplicate_file = st.file_uploader('Upload duplicate.csv or duplicate.xlsx', type=['csv', 'xls', 'xlsx'])
    if duplicate_file:
        upload_duplicate(duplicate_file)

    st.header('View Database Data')

    if st.button('View slot_booking_new.db Data'):
        view_slot_booking_data()

    if st.button('View duplicate.db Data'):
        view_duplicate_data()

    st.header('Export Data to CSV')

    slot_booking_csv = st.button('Export slot_booking_new.db to CSV')
    if slot_booking_csv:
        csv_data = export_slot_booking_to_csv()
        st.download_button(
            label="Download slot_booking_new.csv",
            data=csv_data,
            file_name="slot_booking_new.csv",
            mime="text/csv"
        )

    duplicate_csv = st.button('Export duplicate.db to CSV')
    if duplicate_csv:
        csv_data = export_duplicate_to_csv()
        st.download_button(
            label="Download duplicate.csv",
            data=csv_data,
            file_name="duplicate.csv",
            mime="text/csv"
        )

# Run the app
if __name__ == '__main__':
    main()
