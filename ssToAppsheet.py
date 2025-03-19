import smartsheet
import os
import pandas as pd
import requests
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
SMARTSHEET_API_KEY = os.getenv("SMARTSHEET_API_KEY")
APPSHEET_API_KEY = os.getenv("APPSHEET_API_KEY")
APPSHEET_APP_ID = os.getenv("APPSHEET_APP_ID")
APPSHEET_TABLE_NAME = os.getenv("APPSHEET_TABLE_NAME")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")

# Initialize Smartsheet Client
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_API_KEY)

# Define SERVICE_ACCOUNT_FILE
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")

# Check if it is correctly loaded
if not SERVICE_ACCOUNT_FILE:
    raise ValueError("❌ ERROR: 'SERVICE_ACCOUNT_FILE' is not set in the .env file!")

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

# Function to extract data from Smartsheet
def extract_sheet_data(sheet_id, sheet_name):
    sheet_data = smartsheet_client.Sheets.get_sheet(sheet_id)
    columns = [col.title for col in sheet_data.columns]

    rows = []
    for row in sheet_data.rows:
        row_data = {columns[i]: cell.value if cell.value else '' for i, cell in enumerate(row.cells)}

        # Get attachments
        attachments = smartsheet_client.Sheets.Attachments.list_row_attachments(sheet_id, row.id).data
        attachment_links = []
        for attachment in attachments:
            file_url = attachment.url
            file_name = attachment.name
            file_path = f"./downloads/{file_name}"

            # Download and upload to Drive
            response = requests.get(file_url, headers={"Authorization": f"Bearer {SMARTSHEET_API_KEY}"})
            with open(file_path, "wb") as file:
                file.write(response.content)

            drive_link = upload_to_drive(file_path, file_name)
            attachment_links.append(drive_link)

        row_data["Attachments"] = ", ".join(attachment_links)
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
    print(f"🔄 Processing: {sheet_name}")

    # Extract data
    df = extract_sheet_data(sheet_id, sheet_name)

    # Push to AppSheet Database
    push_to_appsheet(df)

print("🎉 Migration Completed Successfully!")
