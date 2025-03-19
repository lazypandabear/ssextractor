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

# ✅ Load environment variables
load_dotenv()
SMARTSHEET_API_KEY = os.getenv("SMARTSHEET_API_KEY")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")

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

# ✅ Function to Sanitize Sheet Names for Google Sheets
def clean_sheet_name(sheet_name):
    """Removes special characters and replaces spaces for Google Sheets compatibility."""
    sheet_name = re.sub(r"[^\w\s]", "", sheet_name).replace(" ", "_")
    return sheet_name[:30]  # Google Sheets has a 30-character limit

# ✅ Function to Ensure Sheet Tab Exists in Google Sheet
def ensure_google_sheet_tab_exists(google_sheet_id, sheet_name):
    """Ensures the sheet tab exists inside Google Sheets. If not, renames 'Sheet1' or creates a new sheet."""
    sanitized_name = clean_sheet_name(sheet_name)

    # ✅ Get all existing sheets
    sheet_metadata = sheet_service.spreadsheets().get(spreadsheetId=google_sheet_id).execute()
    existing_sheets = [sheet["properties"]["title"] for sheet in sheet_metadata["sheets"]]

    # ✅ If "Sheet1" exists and it's the only sheet, rename it
    if "Sheet1" in existing_sheets and len(existing_sheets) == 1:
        sheet_id = sheet_metadata["sheets"][0]["properties"]["sheetId"]
        rename_request = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {"sheetId": sheet_id, "title": sanitized_name},
                        "fields": "title",
                    }
                }
            ]
        }
        sheet_service.spreadsheets().batchUpdate(
            spreadsheetId=google_sheet_id, body=rename_request
        ).execute()
        print(f"✅ Renamed 'Sheet1' to {sanitized_name}")

    # ✅ If sheet does not exist, create it
    elif sanitized_name not in existing_sheets:
        add_sheet_request = {
            "requests": [{"addSheet": {"properties": {"title": sanitized_name}}}]
        }
        sheet_service.spreadsheets().batchUpdate(
            spreadsheetId=google_sheet_id, body=add_sheet_request
        ).execute()
        print(f"✅ Added new sheet: {sanitized_name}")

    return sanitized_name  # Return the sanitized name for range reference

# ✅ Function to Get or Create a Folder in Google Drive
def get_or_create_drive_folder(folder_name, parent_folder_id):
    """Checks if a folder exists in Google Drive, creates if not, and returns the folder ID."""
    query = f"name='{folder_name}' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
    results = drive_service.files().list(q=query, fields="files(id)").execute()
    
    if results.get("files"):
        return results["files"][0]["id"]  # ✅ Folder already exists, return ID

    # ✅ Create the folder if it doesn't exist
    file_metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id]
    }
    folder = drive_service.files().create(body=file_metadata, fields="id").execute()
    return folder["id"]

# ✅ Function to Ensure Only One Folder Per Sheet in Google Drive
def ensure_sheet_folder(sheet_id):
    """Ensures a single Drive folder exists for each Smartsheet Sheet ID."""
    if sheet_id not in sheet_folders:
        sheet_folders[sheet_id] = get_or_create_drive_folder(str(sheet_id), GOOGLE_DRIVE_FOLDER_ID)
    return sheet_folders[sheet_id]

# ✅ Function to Upload Attachments to a Single Folder Per Sheet
def upload_attachments_to_drive(sheet_id, row_id, file_path, file_name):
    """Uploads attachments to a single Drive folder per Smartsheet sheet ID, with optional row subfolders."""
    
    # ✅ Get or create the main folder for the Smartsheet sheet
    sheet_folder_id = ensure_sheet_folder(sheet_id)
    
    # ✅ Create a sub-folder for each row inside the sheet folder
    row_folder_name = f"Row_{row_id}"
    row_folder_id = get_or_create_drive_folder(row_folder_name, sheet_folder_id)

    # ✅ Upload the file into the row's folder
    file_metadata = {"name": file_name, "parents": [row_folder_id]}
    media = MediaFileUpload(file_path, mimetype="application/octet-stream")
    file = drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()

    return f"https://drive.google.com/file/d/{file.get('id')}/view"

# ✅ Function to Create Google Sheet Based on Smartsheet Sheet ID
def create_google_sheet(sheet_id, sheet_name):
    """Creates a new Google Sheet for the Smartsheet sheet ID and returns its ID."""
    sheet_metadata = {
        "name": str(sheet_id),  # ✅ Use Smartsheet sheet_id as the Google Sheet name
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [GOOGLE_DRIVE_FOLDER_ID]
    }
    
    file = drive_service.files().create(body=sheet_metadata, fields="id").execute()
    google_sheet_id = file.get("id")
    
    print(f"✅ Created Google Sheet: {sheet_id} ({sheet_name}) with ID: {google_sheet_id}")
    return google_sheet_id

# ✅ Function to Extract Data from Smartsheet and Download Attachments
def extract_sheet_data(sheet_id, sheet_name):
    """Extracts data from Smartsheet, downloads attachments, and uploads them to Google Drive."""
    sheet_data = smartsheet_client.Sheets.get_sheet(sheet_id)
    columns = [col.title for col in sheet_data.columns]

    rows = []
    base_dir = f"./downloads/{sheet_id}/"  # ✅ Base directory for the sheet
    os.makedirs(base_dir, exist_ok=True)

    for row in sheet_data.rows:
        row_data = {columns[i]: cell.value if cell.value else '' for i, cell in enumerate(row.cells)}
        row_id = row.id
        row_folder = os.path.join(base_dir, str(row_id))
        os.makedirs(row_folder, exist_ok=True)  # ✅ Ensure row folder exists

        # ✅ Get Attachments for Row
        attachments = smartsheet_client.Attachments.list_row_attachments(sheet_id, row_id).data
        attachment_links = []
        
        for attachment in attachments:
            att_id = attachment.id
            file_name = attachment.name
            file_path = os.path.join(row_folder, file_name)

            # ✅ Fetch Smartsheet Download URL
            retrieve_att = smartsheet_client.Attachments.get_attachment(sheet_id, att_id)
            file_url = retrieve_att.url

            if file_url:
                # ✅ Download File
                response = requests.get(file_url, headers={"Authorization": f"Bearer {SMARTSHEET_API_KEY}"}, stream=True)
                with open(file_path, "wb") as file:
                    for chunk in response.iter_content(chunk_size=8192):
                        file.write(chunk)

                # ✅ Upload to Google Drive & Save Drive Link
                drive_link = upload_attachments_to_drive(sheet_id, row_id, file_path, file_name)
                attachment_links.append(drive_link)
        
        row_data["Attachments"] = ", ".join(attachment_links)  # ✅ Store attachment links
        rows.append(row_data)

    # ✅ Save Data to CSV
    df = pd.DataFrame(rows)
    df.to_csv(f"{sheet_id}.csv", index=False)  # ✅ Use Smartsheet sheet_id as filename
    return df

# ✅ Function to Upload Data to Google Sheets
def upload_to_google_sheets(df, google_sheet_id, sheet_name):
    """Uploads a Pandas DataFrame to Google Sheets, ensuring the sheet tab exists first."""

    sanitized_name = ensure_google_sheet_tab_exists(google_sheet_id, sheet_name)  # ✅ Ensure sheet tab exists
    range_name = f"'{sanitized_name}'!A1"  # ✅ Add single quotes around tab name to prevent parsing issues

    values = [df.columns.tolist()] + df.values.tolist()
    body = {"values": values}

    sheet_service.spreadsheets().values().update(
        spreadsheetId=google_sheet_id,
        range=range_name,
        valueInputOption="RAW",
        body=body
    ).execute()

    print(f"✅ Uploaded {sheet_name} (as {sanitized_name}) to Google Sheets")

# ✅ Get All Sheets from Smartsheet
response = smartsheet_client.Sheets.list_sheets(include_all=True)
sheets = response.data

# ✅ Process Each Sheet
sheet_id_map = {}  # ✅ Dictionary to store Google Sheet IDs mapped to Smartsheet sheet_id

for sheet in sheets:
    sheet_id = sheet.id
    sheet_name = sheet.name
    print(f"🔄 Processing: {sheet_name} (ID: {sheet_id})")

    # ✅ Create a new Google Sheet for this Smartsheet sheet
    google_sheet_id = create_google_sheet(sheet_id, sheet_name)
    sheet_id_map[sheet_id] = google_sheet_id  # ✅ Store mapping

    # ✅ Extract Data from Smartsheet
    df = extract_sheet_data(sheet_id, sheet_name)

    # ✅ Upload to Google Sheets
    upload_to_google_sheets(df, google_sheet_id, sheet_name)

print("🎉 Migration Completed Successfully!")
