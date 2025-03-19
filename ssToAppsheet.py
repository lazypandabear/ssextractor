import smartsheet
import os
import pandas as pd
import requests
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
from dotenv import load_dotenv
import time
import urllib.parse  # For URL encoding

# Load environment variables
load_dotenv()
SMARTSHEET_API_KEY = os.getenv("SMARTSHEET_API_KEY")
APPSHEET_API_KEY = os.getenv("APPSHEET_API_KEY")
APPSHEET_APP_ID = os.getenv("APPSHEET_APP_ID")
APPSHEET_TABLE_NAME = os.getenv("APPSHEET_TABLE_NAME")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")

# Initialize Smartsheet Client
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_API_KEY)

# Function to ensure a directory exists
def ensure_directory(path):
    if not os.path.exists(path):
        os.makedirs(path)

# Define SERVICE_ACCOUNT_FILE
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")

# Check if it is correctly loaded
if not SERVICE_ACCOUNT_FILE:
    raise ValueError("❌ ERROR: 'GOOGLE_SERVICE_ACCOUNT_JSON' is not set in the .env file!")

# Google API Credentials
SCOPES = ["https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = "service_account.json"
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build("drive", "v3", credentials=credentials)

# Function to upload attachments to Google Drive
def upload_to_drive(file_path, file_name):
    file_metadata = {"name": file_name, "parents": [GOOGLE_DRIVE_FOLDER_ID]}
    media = MediaFileUpload(file_path, mimetype="application/octet-stream")
    file = drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    return f"https://drive.google.com/file/d/{file.get('id')}/view"



# Function to fetch Smartsheet attachments and save them
def fetch_attachments(sheet_id, row_id):
    """Fetches and downloads Smartsheet attachments using Smartsheet API (not AWS S3 URLs)."""
    attachment_folder = f"./attachments/{sheet_id}/{row_id}/"
    os.makedirs(attachment_folder, exist_ok=True)

    attachments = smartsheet_client.Attachments.list_row_attachments(sheet_id, row_id).data
    attachment_links = []

    for attachment in attachments:
        attachment_id = attachment.id
        file_name = attachment.name if attachment.name else "unknown_file.txt"
        file_path = os.path.join(attachment_folder, file_name)

        print(f"📥 Fetching file: {file_name}")

        # ✅ Step 1: Get the Smartsheet API download URL (not AWS S3)
        attachment_details = smartsheet_client.Attachments.get_attachment(sheet_id, attachment_id)
        file_url = getattr(attachment_details, 'url', None)

        if not file_url:
            print(f"⚠️ Skipping attachment '{file_name}' (No valid URL found for Sheet {sheet_id}, Row {row_id})")
            continue  # Skip this attachment if URL is missing

        # ✅ Step 2: Use Smartsheet API to Download the File (not AWS S3)
        for attempt in range(3):
            try:
                headers = {
                    "Authorization": f"Bearer {SMARTSHEET_API_KEY}",
                    "Accept": "application/octet-stream"
                }

                # ✅ Stream download to handle large files
                with requests.get(file_url, headers=headers, allow_redirects=True, stream=True, timeout=15) as response:
                    response.raise_for_status()

                    with open(file_path, "wb") as file:
                        for chunk in response.iter_content(chunk_size=8192):
                            file.write(chunk)

                print(f"✅ Successfully downloaded '{file_name}' to {file_path}")
                attachment_links.append(file_path)
                break  # Exit retry loop on success

            except requests.exceptions.RequestException as e:
                print(f"⚠️ Attempt {attempt+1} failed for '{file_name}': {e}")
                time.sleep(2)  # Wait before retrying

        else:
            print(f"❌ Failed to download '{file_name}' after multiple attempts.")

    return attachment_links

# Function to extract Smartsheet Data and process attachments
def extract_sheet_data(sheet_id, sheet_name):
    sheet_data = smartsheet_client.Sheets.get_sheet(sheet_id)
    columns = [col.title for col in sheet_data.columns]

    rows = []
    for row in sheet_data.rows:
        row_data = {columns[i]: cell.value if cell.value else '' for i, cell in enumerate(row.cells)}
        
        # ✅ Fetch and store attachments per row
        row_data["Attachments"] = fetch_attachments(sheet_id, row.id)

        rows.append(row_data)

    df = pd.DataFrame(rows)
    return df



# Function to push data to AppSheet Database
def push_to_appsheet(df):
    APPSHEET_API_URL = f"https://api.appsheet.com/api/v2/apps/{APPSHEET_APP_ID}/tables/{APPSHEET_TABLE_NAME}/Action"

    # Convert DataFrame to JSON payload
    payload = {
        "Action": "Add",
        "Properties": {"Locale": "en-US"},
        "Rows": df.to_dict(orient="records"),
    }

    headers = {"ApplicationAccessKey": APPSHEET_API_KEY, "Content-Type": "application/json"}
    response = requests.post(APPSHEET_API_URL, headers=headers, data=json.dumps(payload))

    if response.status_code == 200:
        print("✅ Successfully pushed data to AppSheet Database")
    else:
        print(f"❌ Failed to push data: {response.status_code} - {response.text}")

# Get all sheets from Smartsheet
response = smartsheet_client.Sheets.list_sheets(include_all=True)
sheets = response.data

# Process each sheet
for sheet in sheets:
    sheet_id = sheet.id
    sheet_name = sheet.name
    print(f"🔄 Processing: {sheet_name} (ID: {sheet_id})")

    # Extract data
    df = extract_sheet_data(sheet_id, sheet_name)

    # Save extracted data
    df.to_csv(f"./attachments/{sheet_id}/sheet_data.csv", index=False)

# Push to AppSheet Database
push_to_appsheet(df)

print("🎉 Migration Completed Successfully!")
