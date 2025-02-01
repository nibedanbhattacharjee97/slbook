import streamlit as st
import pandas as pd
import sqlite3
import datetime
import calendar
import io

# Initialize Database Connection
def init_db():
    conn = sqlite3.connect("app_data.db")
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS appointment_bookings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        manager TEXT,
        spoc TEXT,
        date TEXT,
        slot TEXT
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS studentcap (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        Name TEXT,
        Contact TEXT,
        State TEXT,
        Remarks TEXT
    )''')
    conn.commit()
    conn.close()

# Utility function to insert booking
def insert_booking(manager, spoc, date, slot):
    conn = sqlite3.connect("app_data.db")
    c = conn.cursor()
    c.execute("INSERT INTO appointment_bookings (manager, spoc, date, slot) VALUES (?, ?, ?, ?)", (manager, spoc, date, slot))
    conn.commit()
    conn.close()

# Utility function to check existing booking
def is_slot_available(date, slot):
    conn = sqlite3.connect("app_data.db")
    c = conn.cursor()
    c.execute("SELECT * FROM appointment_bookings WHERE date = ? AND slot = ?", (date, slot))
    result = c.fetchone()
    conn.close()
    return result is None

# UI for slot booking
st.title("Manager-SPOC Slot Booking System")

st.subheader("Book an Appointment")
manager_name = st.text_input("Manager Name")
spoc_name = st.text_input("SPOC Name")

booking_date = st.date_input("Select Date")
selected_slot = st.selectbox("Select Slot", ["10:00 AM - 11:00 AM", "12:00 PM - 1:00 PM", "3:00 PM - 4:00 PM"])

holidays = ["2024-01-01", "2024-08-15", "2024-12-25"]  # Example holiday dates
if str(booking_date) in holidays:
    st.warning("Selected date is a holiday. Please choose another date.")
elif calendar.weekday(booking_date.year, booking_date.month, booking_date.day) in [5, 6]:
    st.warning("Weekends are not allowed for bookings.")
else:
    if st.button("Book Slot"):
        if manager_name and spoc_name:
            if is_slot_available(str(booking_date), selected_slot):
                insert_booking(manager_name, spoc_name, str(booking_date), selected_slot)
                st.success("Appointment booked successfully.")
            else:
                st.error("This slot is already booked.")
        else:
            st.error("Please enter both Manager and SPOC names.")

# Display existing bookings
st.subheader("Existing Bookings")
conn = sqlite3.connect("app_data.db")
bookings_df = pd.read_sql_query("SELECT * FROM appointment_bookings", conn)
conn.close()

if not bookings_df.empty:
    st.dataframe(bookings_df)
else:
    st.info("No bookings found.")

# Student Data Upload Section
st.subheader("Upload Student Data for SPOC Calls")
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if all(col in df.columns for col in ["Name", "Contact", "State", "Remarks"]):
        conn = sqlite3.connect("app_data.db")
        df.to_sql("studentcap", conn, if_exists="replace", index=False)
        conn.close()
        st.success("Data uploaded successfully.")
    else:
        st.error("The Excel file must contain columns: Name, Contact, State, Remarks.")

# Display Student Data
st.subheader("View Student Data")
conn = sqlite3.connect("app_data.db")
student_df = pd.read_sql_query("SELECT * FROM studentcap", conn)
conn.close()

if not student_df.empty:
    st.dataframe(student_df)
else:
    st.info("No student data available.")

# Bulk Delete Student Records
st.subheader("Bulk Delete Student Data")
if st.button("Delete All Student Data"):
    conn = sqlite3.connect("app_data.db")
    conn.execute("DELETE FROM studentcap")
    conn.commit()
    conn.close()
    st.success("All student data deleted.")

# Download Booking Data as CSV
st.subheader("Download Booking Data")
if st.button("Download CSV"):
    conn = sqlite3.connect("app_data.db")
    booking_data = pd.read_sql_query("SELECT * FROM appointment_bookings", conn)
    conn.close()

    csv = booking_data.to_csv(index=False)
    b = io.BytesIO()
    b.write(csv.encode())
    b.seek(0)

    st.download_button(
        label="Download Bookings CSV",
        data=b,
        file_name="bookings.csv",
        mime="text/csv"
    )

# Initialize the database
init_db()
