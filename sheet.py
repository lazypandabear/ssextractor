import smartsheet
import os
import pandas as pd
import requests
import json
import re  # ✅ For sanitizing sheet names
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
from dotenv import load_dotenv
import shutil  # ✅ For file operations

# ✅ Load environment variables
load_dotenv(override=True)
SMARTSHEET_API_KEY = os.getenv("SMARTSHEET_API_KEY")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")

# ✅ Load AppSheet API Key from .env
APPSHEET_API_KEY = os.getenv("APPSHEET_API_KEY")
APPSHEET_APP_ID = os.getenv("APPSHEET_APP_ID")
APPSHEET_TABLE_NAME = os.getenv("APPSHEET_TABLE_NAME")


# ✅ Initialize Smartsheet Client
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_API_KEY)


# ✅ Google API Credentials
SCOPES = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "service_account.json"
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build("drive", "v3", credentials=credentials)
sheet_service = build("sheets", "v4", credentials=credentials)

# ✅ Dictionary to track Google Drive folders for each Smartsheet Sheet ID
sheet_folders = {}

def add_missing_columns_to_appsheet(google_columns, appsheet_columns):
    """Adds missing columns to the AppSheet Table to match Google Sheets."""
    missing_columns = [col for col in google_columns if col not in appsheet_columns]

    if not missing_columns:
        print("✅ No missing columns. AppSheet is up-to-date.")
        return
    
    print(f"➕ Adding missing columns to AppSheet: {missing_columns}")

    url = f"https://api.appsheet.com/api/v2/apps/{APPSHEET_APP_ID}/tables/{APPSHEET_TABLE_NAME}/Schema"
    headers = {
        "Content-Type": "application/json",
        "ApplicationAccessKey": APPSHEET_API_KEY
    }

    # ✅ Construct the payload for adding columns (Defaulting to "Text" type)
    new_columns = [{"Name": col, "Type": "Text"} for col in missing_columns]  
    payload = {"Columns": new_columns}

    # ✅ Send request to update AppSheet schema
    response = requests.patch(url, headers=headers, json=payload)

    if response.status_code == 200:
        print(f"✅ Successfully added missing columns to AppSheet.")
    else:
        print(f"❌ Failed to add columns: {response.text}")


def get_google_sheet_columns(google_sheet_id, sheet_name):
    """Fetches the column names from a Google Sheet."""
    range_name = f"'{sheet_name}'!A1:Z1"  # ✅ First row contains column names
    response = sheet_service.spreadsheets().values().get(
        spreadsheetId=google_sheet_id,
        range=range_name
    ).execute()

    columns = response.get("values", [[]])[0]  # ✅ Extract column names
    print(f"📝 Google Sheets Columns: {columns}")
    return columns

def get_appsheet_columns():
    """Fetches the current column names from AppSheet Database."""
    url = f"https://api.appsheet.com/api/v2/apps/{APPSHEET_APP_ID}/tables/{APPSHEET_TABLE_NAME}/Schema"
    headers = {
        "Content-Type": "application/json",
        "ApplicationAccessKey": APPSHEET_API_KEY
    }

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        schema = response.json()
        columns = [col["Name"] for col in schema.get("Columns", [])]
        print(f"📊 AppSheet Columns: {columns}")
        return columns
    else:
        print(f"❌ Failed to fetch AppSheet columns: {response.text}")
        return []
    
def align_columns_and_upload(df, google_sheet_id, sheet_name):
    """Ensures AppSheet columns match Google Sheets before sending data."""
    
    # ✅ Step 1: Get Google Sheet Columns
    google_columns = get_google_sheet_columns(google_sheet_id, sheet_name)

    # ✅ Step 2: Get AppSheet Columns
    appsheet_columns = get_appsheet_columns()

    # ✅ Step 3: Add Missing Columns
    add_missing_columns_to_appsheet(google_columns, appsheet_columns)

    # ✅ Step 4: Send Data to AppSheet
    send_data_to_appsheet_database(df, google_sheet_id, sheet_name)


def send_data_to_appsheet_database(df, google_sheet_id, sheet_name):
    """Sends extracted Smartsheet data to AppSheet Database via API after ensuring column alignment."""

    # ✅ Get updated Google Sheet Columns
    google_columns = get_google_sheet_columns(google_sheet_id, sheet_name)

    # ✅ Format Data for AppSheet
    records = []
    for _, row in df.iterrows():
        record = {}
        for col in google_columns:
            record[col] = row.get(col, "")

        records.append(record)

    # ✅ Send Data to AppSheet
    url = f"https://api.appsheet.com/api/v2/apps/{APPSHEET_APP_ID}/tables/{APPSHEET_TABLE_NAME}/Action"
    headers = {
        "Content-Type": "application/json",
        "ApplicationAccessKey": APPSHEET_API_KEY
    }

    payload = {
        "Action": "AddOrUpdate",
        "Properties": { "Locale": "en-US" },
        "Rows": records
    }

    print("📤 Sending Data to AppSheet:", json.dumps(payload, indent=2))

    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 200:
        print(f"✅ Successfully synced data with AppSheet Database.")
    else:
        print(f"❌ Failed to sync with AppSheet: {response.text}")

# ✅ Process Each Sheet
for sheet in sheets:
    sheet_id = sheet.id
    sheet_name = sheet.name
    print(f"🔄 Processing: {sheet_name} (ID: {sheet_id})")

    # ✅ Step 1: Create a Google Sheet (if not exists)
    google_sheet_id = create_google_sheet(sheet_id, sheet_name)
    sheet_id_map[sheet_id] = google_sheet_id  # Store mapping

    # ✅ Step 2: Extract Data from Smartsheet
    df = extract_sheet_data(sheet_id, sheet_name)

    # ✅ Step 3: Upload to Google Sheets
    upload_to_google_sheets(df, google_sheet_id, sheet_name)

    # ✅ Step 4: Align AppSheet with Google Sheets & Send Data
    align_columns_and_upload(df, google_sheet_id, sheet_name)

    # ✅ Step 5: Clean up local files
    cleanup_downloads("./downloads/")