import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- GOOGLE SHEETS CONNECTION SETUP ---
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

# --- CACHED FETCHING ---
@st.cache_data(ttl=15)
def fetch_sheet_values(worksheet_name):
    ws = get_worksheet(worksheet_name)
    return ws.get_all_values()

@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path)
    df.rename(columns={'Actual_Manager_Column_Name': 'Manager Name', 'Actual_SPOC_Column_Name': 'SPOC Name'}, inplace=True)
    return df

@st.cache_data(ttl=60)
def load_validation_ids():
    try:
        return pd.read_excel('ids.xlsx')
    except Exception as e:
        st.error(f"Could not read validation file 'ids.xlsx'. Error: {e}")
        return None

def clean_id_series(series):
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

# --- BOOKING FUNCTION ---
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
    
    with st.spinner("Processing your booking..."):
        ws.append_row([next_id, date, time_range, manager, spoc, booked_by])
        st.cache_data.clear()  
        st.session_state['last_action_msg'] = f"✅ Slot booked successfully for {spoc} on {date} ({time_range})!"
        st.rerun()

# --- OPTIMIZED UPLOAD FUNCTION ---
def update_another_database(file):
    with st.spinner("Processing data matching against validation sheet..."):
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

        ws = get_worksheet('plana')
        
        try:
            existing_records_count = len(ws.col_values(1)) - 1
        except Exception:
            existing_records_count = len(ws.get_all_values()) - 1
            
        if existing_records_count < 0:
            existing_records_count = 0
            
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
            
    with st.spinner("Uploading records to Google Sheets..."):
        try:
            ws.append_rows(rows_to_insert, value_input_option='USER_ENTERED')
            st.cache_data.clear() 
            st.session_state['data_uploaded'] = True
            st.session_state['last_action_msg'] = f"✅ Success! {len(filtered_df)} valid student records uploaded and processed successfully."
            st.rerun()
        except Exception as e:
            st.error(f"Upload failed. Google network rejected payload size. Try breaking down your sheets into smaller chunks. Error: {e}")

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

# --- MAIN EXECUTIVE APPLICATION ---
def main():
    st.title('Slot Booking Platform')
    
    # Persistent Success Messaging banner across reruns
    if 'last_action_msg' in st.session_state:
        st.success(st.session_state['last_action_msg'])
        del st.session_state['last_action_msg']

    data = load_data('managers_spocs.xlsx')

    selected_manager = st.selectbox('Select Manager', data['Manager Name'].unique())
    spocs_for_manager = data[data['Manager Name'] == selected_manager]['SPOC Name'].tolist()
    selected_spoc = st.selectbox('Select SPOC', spocs_for_manager)

    selected_date = st.date_input('Select Date')
    time_ranges = ['10:00 AM - 11:00 AM', '11:00 AM - 12:00 PM', '12:00 PM - 1:00 PM', '2:00 PM - 3:00 PM', '3:00 PM - 4:00 PM']
    selected_time_range = st.selectbox('Select Time', time_ranges)
    booked_by = st.text_input('Slot Booked By')

    # --- NEW: LIVE PREVIEW MATRIX ---
    st.markdown("---")
    st.subheader("Current Entry Preview Details")
    preview_df = pd.DataFrame([{
        "Date": str(selected_date),
        "Time Range": selected_time_range,
        "Manager": selected_manager,
        "SPOC": selected_spoc,
        "Booked By": booked_by if booked_by else "⚠️ Required field missing"
    }])
    st.dataframe(preview_df, use_container_width=True, hide_index=True)
    st.markdown("---")

    st.subheader('Upload Student Data For SPOC Calling')
    st.markdown('**Please ensure that only student data marked as Not Joined or Not Contacted from the M&E database is included.**')

    file = st.file_uploader('Upload Excel', type=['xlsx', 'xls'])
    data_uploaded = st.session_state.get('data_uploaded', False)
    
    if file is not None:
        if st.button('Update Data', use_container_width=True):
            update_another_database(file)

    if not data_uploaded:
        st.warning('Please upload student data before booking a slot.')
    else:
        if st.button('Book Slot', type="primary", use_container_width=True):
            insert_booking(str(selected_date), selected_time_range, selected_manager, selected_spoc, booked_by)

    # --- OPTIMIZED CLEAN DOWNLOADING LOGIC ---
    st.subheader('Data Operations & Formats')
    col1, col2, col3 = st.columns(3)
    
    with col1:
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
        df_sample = pd.DataFrame(sample_data)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_sample.to_excel(writer, index=False, sheet_name='Sheet1')
        st.download_button(
            label="📋 Download Sample Format",
            data=output.getvalue(),
            file_name="Sample_Excel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # Fetch fresh bookings tracking matrix
    try:
        all_vals = fetch_sheet_values('slot_booking_new')
        bookings = pd.DataFrame(all_vals[1:], columns=all_vals[0]) if len(all_vals) > 1 else pd.DataFrame()
    except Exception:
        bookings = pd.DataFrame()
    
    if 'date' in bookings.columns and not bookings.empty:
        bookings['date'] = pd.to_datetime(bookings['date'], errors='coerce')

    with col2:
        all_plana = fetch_sheet_values('plana')
        if len(all_plana) > 1:
            df_plana = pd.DataFrame(all_plana[1:], columns=all_plana[0])
            ids_df = load_validation_ids()
            if ids_df is not None and 'cmis_id' in df_plana.columns:
                valid_ids = clean_id_series(ids_df['CMIS_ID']).unique()
                df_plana['cmis_id'] = clean_id_series(df_plana['cmis_id'])
                filtered_plana = df_plana[df_plana['cmis_id'].isin(valid_ids)]
                
                st.download_button(
                    label="📥 Download M&E Verified Data",
                    data=filtered_plana.to_csv(index=False),
                    file_name="plana_filtered.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        else:
            st.button("📥 Download M&E Verified Data", disabled=True, use_container_width=True)

    with col3:
        if not bookings.empty:
            st.download_button(
                label="🗓️ Download Monthly Bookings",
                data=bookings.to_csv(index=False),
                file_name="monthly_bookings.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.button("🗓️ Download Monthly Bookings", disabled=True, use_container_width=True)

    st.subheader('Calendar View (Current Month Status)')
    st.markdown(generate_calendar(bookings), unsafe_allow_html=True)

    st.header("Today's Bookings")
    current_date = datetime.now().strftime("%Y-%m-%d")

    if not bookings.empty and 'date' in bookings.columns:
        today_bookings_df = bookings[bookings['date'].dt.strftime("%Y-%m-%d") == current_date]
        if not today_bookings_df.empty:
            st.write(f"Bookings for today ({current_date}):")
            for _, row in today_bookings_df.iterrows():
                st.write(f"- **Time Slot:** {row.get('time_range', '')} | **Manager:** {row.get('manager', '')} | **SPOC:** {row.get('spoc', '')}")
        else:
            st.info("No bookings scheduled for today.")
    else:
        st.info("No bookings scheduled for today.")

if __name__ == '__main__':
    main()