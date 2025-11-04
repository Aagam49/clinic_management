import os
import ssl
import certifi
import time
from datetime import datetime
from typing import List
from flask import Flask, request, jsonify, render_template
from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv

# ------------------- Load environment -------------------
load_dotenv()
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")
SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME")

if not GOOGLE_CREDENTIALS_FILE or not os.path.exists(GOOGLE_CREDENTIALS_FILE):
    raise FileNotFoundError("Google service account file not found. Set GOOGLE_APPLICATION_CREDENTIALS in .env")
if not SHEET_ID:
    raise EnvironmentError("SHEET_ID not set in environment.")
if not SHEET_NAME:
    raise EnvironmentError("SHEET_NAME not set in environment.")

# ------------------- SSL Fix -------------------
ssl._create_default_https_context = lambda: ssl.create_default_context(cafile=certifi.where())

# ------------------- Google Sheets Setup -------------------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
_SERVICE_ACCOUNT_CREDS = service_account.Credentials.from_service_account_file(
    GOOGLE_CREDENTIALS_FILE, scopes=SCOPES
)

def get_sheets_service():
    return build("sheets", "v4", credentials=_SERVICE_ACCOUNT_CREDS)

# ------------------- Cache System -------------------
CACHE_TTL = 30
_last_cache_time = 0
_cached_patients = []

# ------------------- Canonical sheet headers -------------------
# IMPORTANT: This list now reflects the order used in your latest form and Google Sheet headers.
CANONICAL_HEADERS = [
    "Patient ID",
    "Name",
    "Number",
    "Age",
    "Gender",
    "Occupation",
    "Ref. by",
    "Address",
    "Date of joining",
    "Conditions",
    "Time",
    "Visit Days",
    "Visit Count",
]

# For robustness, map canonical -> accepted input keys (variants)
ALLOWED_KEY_VARIANTS = {
    "Patient ID": ["Patient ID", "Patient_ID", "patient_id", "patient id"],
    "Name": ["Name", "name"],
    "Number": ["Number", "number"],
    "Age": ["Age", "age"],
    "Gender": ["Gender", "gender"],
    "Occupation": ["Occupation", "occupation"],
    "Ref. by": ["Ref. by", "Ref.by", "Ref_by", "Ref by", "ref by", "ref.by"],
    "Address": ["Address", "address"],
    "Date of joining": ["Date of joining", "Date_of_joining", "Date of Joining", "date of joining"],
    "Conditions": ["Conditions", "conditions"],
    "Time": ["Time", "time"],
    "Visit Days": ["Visit Days", "Visit_Days", "visit days", "visit_days"],
    "Visit Count": ["Visit Count", "Visit_Count", "visit_count"],
}

def get_cached_patients():
    """
    Load patients from Google Sheets with 30s cache.
    Uses canonical headers to enforce structure, improving robustness.
    """
    global _last_cache_time, _cached_patients
    now = time.time()
    if now - _last_cache_time > CACHE_TTL:
        print("🔄 Refreshing from Google Sheets...")
        service = get_sheets_service()
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID, range=SHEET_NAME
        ).execute()
        values = result.get('values', [])
        
        if not values:
            _cached_patients = []
        else:
            headers_from_sheet = values[0]
            data_rows = values[1:]
            _cached_patients = []
            
            for row in data_rows:
                patient_dict = {}
                for canonical_h in CANONICAL_HEADERS:
                    try:
                        # Find the index of the canonical header in the actual sheet headers
                        sheet_col_index = headers_from_sheet.index(canonical_h)
                        # Use the value from the current row at that index, or "" if the row is too short
                        patient_dict[canonical_h] = row[sheet_col_index] if sheet_col_index < len(row) else ""
                    except ValueError:
                        # If a canonical header is not in the sheet, set to empty
                        patient_dict[canonical_h] = ""
                        
                _cached_patients.append(patient_dict)
                
        _last_cache_time = now
    return _cached_patients

# ------------------- Helpers -------------------
def column_to_letter(col_index: int) -> str:
    letters = ""
    while col_index >= 0:
        letters = chr((col_index % 26) + ord("A")) + letters
        col_index = col_index // 26 - 1
    return letters

def parse_visit_days(raw: str) -> List[str]:
    if not raw:
        return []
    s = str(raw).lower()
    for sep in [";", "/", "\\", "|", " - ", "-", " to "]:
        s = s.replace(sep, ",")
    s = s.replace("\n", ",")
    tokens = [t.strip() for t in s.split(",") if t.strip()]
    return tokens

def matches_today(visit_tokens: List[str]) -> bool:
    today_full = datetime.today().strftime("%A").lower()
    today_short = today_full[:3]
    for t in visit_tokens:
        if "daily" in t:
            return True
        if t == today_full or t == today_short:
            return True
        if today_full in t or today_short in t:
            return True
    return False

def find_value_for_header(data_dict: dict, header: str):
    """
    Given the incoming JSON (data_dict) and a canonical header,
    try all allowed variants and also try case-insensitive matches.
    """
    variants = ALLOWED_KEY_VARIANTS.get(header, [header])
    # direct check
    for k in variants:
        if k in data_dict:
            return data_dict.get(k)
    # try case-insensitive keys
    lower_map = {str(k).lower(): v for k, v in data_dict.items()}
    for k in variants:
        if str(k).lower() in lower_map:
            return lower_map[str(k).lower()]
    # finally try header itself (exact)
    if header in data_dict:
        return data_dict.get(header)
    return ""


# ------------------- Flask App -------------------
app = Flask(__name__)

# ---- API Routes ----
@app.route("/api/patients", methods=["GET"])
def get_patients():
    try:
        patients = get_cached_patients()
        return jsonify(patients)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/patients/today", methods=["GET"])
def get_today_patients():
    try:
        patients = get_cached_patients()
        today_patients = []
        for p in patients:
            tokens = parse_visit_days(p.get("Visit Days", ""))
            if matches_today(tokens):
                p.setdefault("Visit Count", "0")
                p.setdefault("Patient ID", "")
                p.setdefault("Patient_ID", "")
                today_patients.append(p)
        return jsonify(today_patients)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/patients", methods=["POST"])
def add_patient():
    try:
        data = request.get_json(force=True) or {}
        print("Incoming add_patient payload:", data)
        
        # --- Prepare Data ---
        # 1. Handle Visit Days (sent as a list from the form)
        raw_days = find_value_for_header(data, "Visit Days") or ""
        if isinstance(raw_days, list):
            visit_days_str = ", ".join([str(x).strip().capitalize() for x in raw_days if str(x).strip()])
        else:
            visit_days_str = str(raw_days).strip()
            
        # 2. Map payload data to a standardized dictionary
        standardized_data = {}
        for canonical_h in CANONICAL_HEADERS:
            standardized_data[canonical_h] = find_value_for_header(data, canonical_h) or ""
            
        # 3. Apply special processing (Visit Days, Visit Count)
        standardized_data["Visit Days"] = visit_days_str
        standardized_data["Visit Count"] = str(standardized_data.get("Visit Count", "0"))

        # --- Build new row strictly by canonical headers ---
        # This list comprehension ensures the final row is always in the exact canonical order.
        new_row = [standardized_data.get(header, "") for header in CANONICAL_HEADERS]
        
        # Append to sheet
        service = get_sheets_service()
        service.spreadsheets().values().append(
            spreadsheetId=SHEET_ID,
            range=SHEET_NAME,
            valueInputOption="USER_ENTERED", 
            body={"values": [new_row]},
        ).execute()

        global _last_cache_time
        _last_cache_time = 0

        return jsonify({"status": "success"}), 201

    except Exception as e:
        print("Error in add_patient:", e)
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/patients/<patient_id>/attend", methods=["PUT"])
def mark_attendance(patient_id):
    try:
        payload = request.get_json(force=True) or {}
        if str(payload.get("action", "")).lower() != "confirm":
            return jsonify({"status": "ignored", "message": "action not confirm"}), 200

        service = get_sheets_service()
        result = service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=SHEET_NAME).execute()
        values = result.get("values", [])
        if not values or len(values) < 2:
            return jsonify({"status": "error", "message": "No data in sheet"}), 404

        headers = values[0]
        data_rows = values[1:]
        # find the column indices robustly (try variants)
        try:
            pid_col = headers.index("Patient ID")
        except ValueError:
            # fallback to Patient_ID
            pid_col = headers.index("Patient_ID") if "Patient_ID" in headers else None

        try:
            visit_count_col = headers.index("Visit Count")
        except ValueError:
            visit_count_col = headers.index("Visit_Count") if "Visit_Count" in headers else None

        if pid_col is None or visit_count_col is None:
            return jsonify({"status": "error", "message": "Required columns missing in sheet"}), 500

        updated = False
        for row_idx, row in enumerate(data_rows, start=2):
            current_pid = str(row[pid_col]) if len(row) > pid_col else ""
            if current_pid == str(patient_id):
                current_count = int(row[visit_count_col] or "0") if len(row) > visit_count_col else 0
                new_count = current_count + 1
                col_letter = column_to_letter(visit_count_col)
                range_to_update = f"{SHEET_NAME}!{col_letter}{row_idx}"
                service.spreadsheets().values().update(
                    spreadsheetId=SHEET_ID,
                    range=range_to_update,
                    valueInputOption="USER_ENTERED",
                    body={"values": [[str(new_count)]]},
                ).execute()
                updated = True
                break

        if updated:
            return jsonify({"status": "updated", "patient_id": patient_id, "new_count": new_count}), 200
        else:
            return jsonify({"status": "not found", "patient_id": patient_id}), 404
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

# ---- Web Routes ----
@app.route('/')
def index():
    patients = get_cached_patients()
    today_patients = [p for p in patients if matches_today(parse_visit_days(p.get("Visit Days", "")))]
    return render_template('index.html',
                           total_count=len(patients),
                           today_count=len(today_patients))

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
