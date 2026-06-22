import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
import base64
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- GOOGLE SHEETS CONNECTION SETUP ---
# Cached connection client so it doesn't re-authenticate on every interaction
@st.cache_resource
def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    try:
        creds_dict = {
            "type": st.secrets["connections"]["gsheets"]["type"],
            "project_id": st.secrets["connections"]["gsheets"]["project_id"],
            "private_key_id": st.secrets["connections"]["gsheets"]["private_key_id"],
            "private_key": st.secrets["connections"]["gsheets"]["private_key"],
            "client_email": st.secrets["connections"]["gsheets"]["client_email"],
            "client_id": st.secrets["connections"]["gsheets"]["client_id"],
            "auth_uri": st.secrets["connections"]["gsheets"]["auth_uri"],
            "token_uri": st.secrets["connections"]["gsheets"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["connections"]["gsheets"]["auth_provider_x509_cert_url"],
            "client_x509_cert_url": st.secrets["connections"]["gsheets"]["client_x509_cert_url"]
        }
    except KeyError:
        st.error("### ❌ Secrets Configuration Error\n"
                 "Streamlit could not find the `[connections.gsheets]` structure in your secrets.\n\n"
                 "Please verify that your configuration file matches the format you provided.")
        st.stop()
    
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client

def get_worksheet(worksheet_name):
    try:
        client = get_gspread_client()
        spreadsheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
        sheet = client.open_by_url(spreadsheet_url)
    except Exception as e:
        st.error(f"### ❌ Connection Error\nCould not access the Google Spreadsheet URL. Details: {e}")
        st.stop()
    
    try:
        return sheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        if worksheet_name == 'slot_booking_new':
            ws = sheet.add_worksheet(title='slot_booking_new', rows="1000", cols="20")
            ws.append_row(["id", "date", "time_range", "manager", "spoc", "booked_by"])
            return ws
        elif worksheet_name == 'plana':
            ws = sheet.add_worksheet(title='plana', rows="5000", cols="20")
            ws.append_row(["id", "cmis_id", "student_name", "cmis_ph_no", "center_name", 
                           "uploader_name", "verification_type", "mode_of_verification", "verification_date"])
            return ws
        else:
            raise

# --- OPTIMIZATION: Cached Fetching Functions ---
# We cache these pulls for 10 seconds to stop it pulling from Google on every widget toggle
@st.cache_data(ttl=10)
def fetch_sheet_values(worksheet_name):
    ws = get_worksheet(worksheet_name)
    return ws.get_all_values()

@st.cache_data
def load_data(file_path):
    # If it's a fixed file name string, read it once
    df = pd.read_excel(file_path)
    df.rename(columns={'Actual_Manager_Column_Name': 'Manager Name', 'Actual_SPOC_Column_Name': 'SPOC Name'}, inplace=True)
    return df

@st.cache_data
def load_validation_ids():
    try:
        return pd.read_excel('ids.xlsx')
    except Exception as e:
        st.error(f"Could not read validation file 'ids.xlsx'. Error: {e}")
        return None

def clean_id_series(series):
    """Helper function to clean float representations and extra whitespaces from IDs"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

def insert_booking(date, time_range, manager, spoc, booked_by):
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

    # Use cached fetch instead of direct API call
    all_vals = fetch_sheet_values('slot_booking_new')
    headers = all_vals[0] if all_vals else []
    records = [dict(zip(headers, row)) for row in all_vals[1:]] if len(all_vals) > 1 else []
    
    existing_booking = any(
        str(r.get('date', '')) == str(date) and str(r.get('spoc', '')).strip().lower() == str(spoc).strip().lower() 
        for r in records
    )

    if existing_booking:
        st.error('Slot booking failed. This SPOC is already booked for the selected date.')
        return

    next_id = len(all_vals)
    ws = get_worksheet('slot_booking_new')
    ws.append_row([next_id, date, time_range, manager, spoc, booked_by])
    
    # Clear cache so next screen load reflects the updated database instantly
    st.cache_data.clear()
    st.success('Slot booked successfully!')
    st.rerun()

def update_another_database(file):
    df = pd.read_excel(file)
    ids_df = load_validation_ids()
    
    if ids_df is None:
        return

    valid_ids = clean_id_series(ids_df['CMIS_ID']).unique()
    df['CMIS ID'] = clean_id_series(df['CMIS ID'])
    filtered_df = df[df['CMIS ID'].isin(valid_ids)]

    if filtered_df.empty:
        st.error("No valid records matched the validation IDs. Check if the IDs in your uploaded sheet match 'ids.xlsx'.")
        return

    # Use cached call to grab data footprint size safely
    all_vals = fetch_sheet_values('plana')
    existing_records_count = max(0, len(all_vals) - 1)
        
    rows_to_insert = []
    for index, row in filtered_df.iterrows():
        existing_records_count += 1
        rows_to_insert.append([
            existing_records_count, 
            str(row.get('CMIS ID', '')), 
            str(row.get('Student Name', '')), 
            str(row.get('CMIS PH No(10 Number)', '')),
            str(row.get('Center Name', '')), 
            str(row.get('Name Of Uploder', '')), 
            str(row.get('Verification Type', '')), 
            str(row.get('Mode Of Verification', '')), 
            str(row.get('Verification Date', ''))
        ])
        
    ws = get_worksheet('plana')
    ws.append_rows(rows_to_insert)
    
    # Clear cache to reset local state data
    st.cache_data.clear()
    st.success(f"{len(filtered_df)} valid records inserted successfully into Google Sheets.")
    st.rerun()

def download_another_database_data():
    all_vals = fetch_sheet_values('plana')
    
    if len(all_vals) <= 1:
        st.error("No valid data found for M&E verification.")
        return
        
    headers = all_vals[0]
    df = pd.DataFrame(all_vals[1:], columns=headers)
    
    ids_df = load_validation_ids()
    if ids_df is None:
        return
        
    valid_ids = clean_id_series(ids_df['CMIS_ID']).unique()
    
    if 'cmis_id' in df.columns:
        df['cmis_id'] = clean_id_series(df['cmis_id'])
        filtered_df = df[df['cmis_id'].isin(valid_ids)]
    else:
        st.error("Missing expected 'cmis_id' structure in Google Sheets database header values.")
        return

    if filtered_df.empty:
        st.error("No valid data found matching your validation IDs.")
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
            
            if not bookings.empty and 'date' in bookings.columns:
                bookings_on_day = bookings[(bookings['date'].dt.year == current_year) &
                                           (bookings['date'].dt.month == current_month) &
                                           (bookings['date'].dt.day == day)]
            else:
                bookings_on_day = pd.DataFrame()
                
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
    
    data = load_data('managers_spocs.xlsx')

    selected_manager = st.selectbox('Select Manager', data['Manager Name'].unique())
    spocs_for_manager = data[data['Manager Name'] == selected_manager]['SPOC Name'].tolist()
    selected_spoc = st.selectbox('Select SPOC', spocs_for_manager)

    selected_date = st.date_input('Select Date')
    time_ranges = ['10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM', '12:00 PM - 1:00 PM', '2:00 PM - 3:00 PM', '3:00 PM - 4:00 PM']
    selected_time_range = st.selectbox('Select Time', time_ranges)
    booked_by = st.text_input('Slot Booked By')

    st.subheader('Upload Student Data For SPOC Calling')
    st.markdown('**Please ensure that only student data marked as Not Joined or Not Contacted from the M&E database is included.**')

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

    try:
        # Optimized with cache fetching function
        all_vals = fetch_sheet_values('slot_booking_new')
        if len(all_vals) > 1:
            bookings = pd.DataFrame(all_vals[1:], columns=all_vals[0])
        else:
            bookings = pd.DataFrame()
    except Exception:
        bookings = pd.DataFrame()
    
    if 'date' in bookings.columns and not bookings.empty:
        bookings['date'] = pd.to_datetime(bookings['date'], errors='coerce')

    st.subheader('Calendar View (Current Month Status)')
    st.markdown(generate_calendar(bookings), unsafe_allow_html=True)

    st.header("Today's Bookings")
    current_date = datetime.now().strftime("%Y-%m-%d")

    if not bookings.empty and 'date' in bookings.columns:
        today_bookings_df = bookings[bookings['date'].dt.strftime("%Y-%m-%d") == current_date]
        if not today_bookings_df.empty:
            st.write(f"Bookings for today ({current_date}):")
            for _, row in today_bookings_df.iterrows():
                st.write(f"- Time Slot: {row.get('time_range', '')}, Manager: {row.get('manager', '')}, SPOC: {row.get('spoc', '')}")
        else:
            st.write("No bookings for today.")
    else:
        st.write("No bookings for today.")

    if st.button('Download Monthly Data'):
        if not bookings.empty:
            csv = bookings.to_csv(index=False)
            b64 = base64.b64encode(csv.encode()).decode()
            href = f'<a href="data:file/csv;base64,{b64}" download="monthly_bookings.csv">Download CSV</a>'
            st.markdown(href, unsafe_allow_html=True)
        else:
            st.warning("No booking data available to extract.")

if __name__ == '__main__':
    main()