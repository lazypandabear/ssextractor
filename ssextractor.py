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
JIRA_API_KEY = os.getenv("JIRA_API_KEY")
JIRA_EMAIL = os.getenv("JIRA_EMAIL")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")

# Initialize Smartsheet Client
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_API_KEY)

# Google API Credentials
SCOPES = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "service_account.json"
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build("drive", "v3", credentials=credentials)
sheet_service = build("sheets", "v4", credentials=credentials)

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
    df.to_csv(f"{sheet_name}.csv", index=False)
    return df

# Function to upload data to Google Sheets
def upload_to_google_sheets(df, sheet_name):
    range_name = f"{sheet_name}!A1"
    values = [df.columns.tolist()] + df.values.tolist()
    body = {"values": values}

    sheet_service.spreadsheets().values().update(
        spreadsheetId=GOOGLE_SHEET_ID,
        range=range_name,
        valueInputOption="RAW",
        body=body
    ).execute()
    print(f"✅ Uploaded {sheet_name} to Google Sheets")

# Function to create Jira issues (optional)
def create_jira_issue(summary, description):
    JIRA_URL = "https://yourcompany.atlassian.net"
    url = f"{JIRA_URL}/rest/api/3/issue"
    headers = {
        "Authorization": f"Basic {JIRA_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "fields": {
            "project": {"key": "YOUR_PROJECT_KEY"},
            "summary": summary,
            "description": description,
            "issuetype": {"name": "Task"}
        }
    }
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    return response.json()

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

    # Upload to Google Sheets
    upload_to_google_sheets(df, sheet_name)

    # Optionally create Jira issues
    for _, row in df.iterrows():
        if "Task Name" in row and "Description" in row:
            issue = create_jira_issue(row["Task Name"], row["Description"])
            print(f"✅ Created Jira Issue: {issue}")

print("🎉 Migration Completed Successfully!")
