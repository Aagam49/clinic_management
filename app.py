import os
import ssl
import certifi
import string
from datetime import datetime
from flask import Flask, request, jsonify, render_template
from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv

# ------------------- Load Environment -------------------
load_dotenv()

# ------------------- Google Sheets Setup -------------------
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")
SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME")

print("Service account file path:", GOOGLE_CREDENTIALS_FILE)
if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
    raise FileNotFoundError("⚠️ Google service account file not found. Check your .env path.")
print("Google Sheet ID:", SHEET_ID)
print("Worksheet name:", SHEET_NAME)

# Fix SSL verification issues on Windows
ssl._create_default_https_context = lambda: ssl.create_default_context(cafile=certifi.where())

# Load service account credentials
creds = service_account.Credentials.from_service_account_file(
    GOOGLE_CREDENTIALS_FILE,
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)

# Build the Sheets API service
service = build("sheets", "v4", credentials=creds)
sheet = service.spreadsheets()

# ------------------- Flask App -------------------
app = Flask(__name__)


# ------------------- Helper Function -------------------
def column_to_letter(col_index: int) -> str:
    """Convert 0-based column index to A1 letter notation."""
    letter = ''
    while col_index >= 0:
        remainder = col_index % 26
        letter = string.ascii_uppercase[remainder] + letter
        col_index = col_index // 26 - 1
    return letter

# ------------------- API ROUTES -------------------

#  Get all patients
@app.route('/api/patients', methods=['GET'])
def get_patients():
    try:
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=SHEET_NAME).execute()
        values = result.get('values', [])
        if not values:
            return jsonify([])

        headers = values[0]
        patients = [dict(zip(headers, row)) for row in values[1:]]
        return jsonify(patients)
    except Exception as e:
        print(f"Error getting all patients: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ✅ Get today's patients
@app.route('/api/patients/today', methods=['GET'])
def get_today_patients():
    try:
        # Fetch full sheet data
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=SHEET_NAME).execute()
        values = result.get('values', [])
        if not values:
            print("⚠️ No data found in sheet.")
            return jsonify([])

        headers = values[0]
        patients = [dict(zip(headers, row)) for row in values[1:]]

        # Determine today's day name
        today_full = datetime.today().strftime('%A').lower()  # e.g., 'saturday'
        today_short = today_full[:3]

        today_patients = []

        print("---- DEBUG VISIT DAYS COLUMN ----")
        for p in patients:
            raw_visit_days = p.get('Visit Days', '')
            visit_days_str = (raw_visit_days or '').lower().replace('\n', ' ').replace(';', ',').replace('/', ',')
            visit_days = [d.strip() for d in visit_days_str.split(',') if d.strip()]

            print(f"Row: {p.get('Name', 'Unknown')} | Visit Days: {visit_days}")

            # Match flexible day formats
            if any(
                today_full.startswith(v[:3]) or
                v.startswith(today_short) or
                today_full in v or
                today_short in v or
                'daily' in v
                for v in visit_days
            ):
                today_patients.append(p)

        print(f"✅ Today's patients count: {len(today_patients)}")
        for tp in today_patients:
            print(f" - {tp.get('Name', 'Unknown')}")

        return jsonify(today_patients)

    except Exception as e:
        import traceback
        print("❌ Error loading today's patients:", e)
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ✅ Add a new patient
@app.route('/api/patients', methods=['POST'])
def add_patient():
    data = request.json or {}
    raw_days = data.get('Visit Days', [])
    visit_days_clean = [v.strip().capitalize() for v in raw_days if v.strip()]

    row = [
        data.get('Patient_ID', ''),
        data.get('Name', ''),
        data.get('Number', ''),
        data.get('Age', ''),
        data.get('Gender', ''),
        data.get('Occupation', ''),
        data.get('Ref.by', ''),
        data.get('Address', ''),
        data.get('Date of joining', ''),
        data.get('Conditions', ''),
        data.get('Time', ''),  # fixed: remove accidental trailing space
        ', '.join(visit_days_clean),
        data.get('Visit Count', '0'),
    ]

    try:
        sheet.values().append(
            spreadsheetId=SHEET_ID,
            range=SHEET_NAME,
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values': [row]}
        ).execute()
        return jsonify({'status': 'success'})
    except Exception as e:
        print(f"Error adding patient: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ✅ Mark attendance / increment Visit Count
@app.route('/api/patients/<patient_id>/attend', methods=['PUT'])
def mark_attendance(patient_id):
    data = request.json or {}
    if data.get('action', '').lower() != 'confirm':
        return jsonify({'status': 'ignored'})

    try:
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=SHEET_NAME).execute()
        values = result.get('values', [])
        if not values:
            return jsonify({'status': 'error', 'message': 'No data in sheet'}), 404

        headers = values[0]
        data_rows = values[1:]

        patient_id_col_index = headers.index('Patient_ID')
        visit_count_col_index = headers.index('Visit Count')

        updated = False
        for i, row in enumerate(data_rows, start=2):
            current_id = str(row[patient_id_col_index]) if len(row) > patient_id_col_index else None
            if current_id == patient_id:
                current_count_str = row[visit_count_col_index] if len(row) > visit_count_col_index else '0'
                current_count = int(current_count_str or '0')
                new_count = current_count + 1

                range_to_update = f"{SHEET_NAME}!{column_to_letter(visit_count_col_index)}{i}"
                sheet.values().update(
                    spreadsheetId=SHEET_ID,
                    range=range_to_update,
                    valueInputOption='USER_ENTERED',
                    body={'values': [[new_count]]}
                ).execute()
                updated = True
                break

        return jsonify({'status': 'updated' if updated else 'not found'})
    except Exception as e:
        print(f"Error in mark_attendance: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ------------------- Frontend Routes -------------------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/today')
def today_page():
    return render_template('today_patients.html')

@app.route('/add')
def add_patient_page():
    return render_template('add_patients.html')

@app.route('/history')
def history_page():
    return render_template('all_patients.html')

# ------------------- Run App -------------------
if __name__ == '__main__':
    app.run(host='127.0.0.1', port=int(os.environ.get('PORT', 5000)), debug=False)
