import os
import pandas as pd
import requests
import glob  # ✅ Used for wildcard search
import smartsheet
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
from dotenv import load_dotenv
import time  # ✅ For sleep

# ✅ Load environment variables
load_dotenv(override=True)
SMARTSHEET_API_KEY = os.getenv("SMARTSHEET_API_KEY")
GOOGLE_DRIVE_SHEETS_FOLDER_ID = os.getenv("GOOGLE_DRIVE_SHEETS_FOLDER_ID")
GOOGLE_DRIVE__COMMENTS_FOLDER_ID = os.getenv("GOOGLE_DRIVE__COMMENTS_FOLDER_ID")
GOOGLE_DRIVE_ATTACHMENTS_FOLDER_ID = os.getenv("GOOGLE_DRIVE_ATTACHMENTS_FOLDER_ID")
APPSHEET_API_KEY = os.getenv("APPSHEET_API_KEY")
APPSHEET_APP_ID = os.getenv("APPSHEET_APP_ID")
APPSHEET_TABLE_NAME = os.getenv("APPSHEET_TABLE_NAME")


# ✅ Google API Credentials
SCOPES = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "service_account.json"
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build("drive", "v3", credentials=credentials)
sheet_service = build("sheets", "v4", credentials=credentials)
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_API_KEY)


# ✅ Ensure folder exists
def ensure_folder(folder_path):
    """Ensures a folder exists before saving files."""
    os.makedirs(folder_path, exist_ok=True)

def download_smartsheet_as_excel(sheet_id):
    """Downloads a Smartsheet as an Excel file and adds the Row ID column efficiently."""
    try:
        # ✅ Define folders and paths
        sheet_folder = os.path.abspath(f"sheets/{sheet_id}")
        os.makedirs(sheet_folder, exist_ok=True)  # Ensure directory exists

        # ✅ Download Excel and save it
        excel_data = smartsheet_client.Sheets.get_sheet_as_excel(sheet_id, sheet_folder)
        excel_data.save_to_file()  # Save file in the directory

# ✅ Find the saved Excel file using wildcard (*.xlsx)
        retry_count = 5
        excel_files = []
        while retry_count > 0:
            excel_files = glob.glob(os.path.join(sheet_folder, "*.xlsx"))
            if excel_files:
                break
            print("⚠️ Waiting for Excel file to be available...")
            time.sleep(2)
            retry_count -= 1

        if not excel_files:
            print(f"❌ No Excel file found in {sheet_folder} after download.")
            return None

        original_file = excel_files[0]  # Use the first (and only) file

        # ✅ Load Excel into Pandas Safely
        with pd.ExcelFile(original_file, engine="openpyxl") as xls:
            sheet_name = xls.sheet_names[0]  # Assume first sheet contains data
            df = pd.read_excel(xls, sheet_name=sheet_name)

        # ✅ Fetch Smartsheet Row IDs in bulk (Efficient)
        sheet_data = smartsheet_client.Sheets.get_sheet(sheet_id)

        row_ids = {row.row_number: row.id for row in sheet_data.rows}  # Map row_number → row_id

        # ✅ Add Row ID column (Efficient mapping)
        df.insert(0, "Row ID", pd.Series(range(1, len(df) + 1)).map(row_ids))

        # ✅ Add "Filename" column
        df["Filename"] = os.path.basename(original_file)

        # ✅ Save the updated file
        updated_excel_path = os.path.join(sheet_folder, f"{sheet_id}.xlsx")
        df.to_excel(updated_excel_path, index=False)

        
       

        print(f"✅ Smartsheet {sheet_id} downloaded and updated with Row ID: {updated_excel_path}")
        return updated_excel_path,original_file

    except Exception as e:
        print(f"❌ Error downloading Smartsheet {sheet_id}: {e}")
        return None
    




def fetch_smartsheet_row_ids(sheet_id):
    """Fetches all row IDs from Smartsheet and returns a row number to row ID mapping."""
    try:
        sheet_data = smartsheet_client.Sheets.get_sheet(sheet_id)
        row_mapping = {row.row_number: row.id for row in sheet_data.rows}  # ✅ Map row number → row ID

        print(f"✅ Retrieved {len(row_mapping)} Smartsheet row IDs for Sheet {sheet_id}")
        return row_mapping

    except Exception as e:
        print(f"❌ Error fetching Smartsheet row IDs for {sheet_id}: {e}")
        return {}
    

# ✅ Extract & Store Comments
def extract_and_store_comments(sheet_id,original_file):
    """Reads Smartsheet Excel, extracts comments, and stores them row-wise."""
    try:
        # ✅ Find the downloaded Excel file
        sheet_folder = os.path.abspath(f"sheets/{sheet_id}")
        excel_files = glob.glob(os.path.join(sheet_folder, "*.xlsx"))
        if not excel_files:
            print(f"❌ Smartsheet Excel not found in {sheet_folder}")
            return

        excel_path = excel_files[0]
        comments_folder = os.path.abspath(f"comments/{sheet_id}")
        ensure_folder(comments_folder)

        # ✅ Find the saved Excel file using wildcard (*.xlsx)
        # ✅ Read Comments sheet
        retry_count = 5
        excel_files = []
        while retry_count > 0:
            excel_files = glob.glob(os.path.join(sheet_folder, original_file))
            if excel_files:
                break
            print("⚠️ Waiting for Excel file to be available...")
            time.sleep(2)
            retry_count -= 1

        if not excel_files:
            print(f"❌ No Excel file found in {sheet_folder} after download.")
            return None
        
        # ✅ Load Excel into Pandas Safely
        with pd.ExcelFile(original_file, engine="openpyxl") as xls:
            df_comments = pd.read_excel(xls, sheet_name="Comments", header=None)

        # ✅ Dynamically assign headers (Fixes length mismatch error)
        expected_columns = ["Relative Row", "Comments", "Created By", "Created On", "Actual Row ID"]
        df_comments = df_comments.iloc[:, :len(expected_columns)]  # Trim extra columns
        df_comments.columns = expected_columns[:df_comments.shape[1]]  # Assign only existing columns
        df_comments = df_comments.dropna(how='all')
        df_comments['Relative Row']= df_comments['Relative Row'].ffill()
        df_comments.to_excel(f"{comments_folder}/{sheet_id}_comments.xlsx", index=False)



        print(f"✅ Saved comments to {comments_folder}/{sheet_id}_comments.xlsx")
        

    except Exception as e:
        print(f"❌ Error extracting comments for Sheet {sheet_id}: {e}")



def create_relative_row_mapping(sheet_id):
    """Creates a mapping table of 'Relative Row' to 'Actual Row ID' from Smartsheet comments data."""
    try:
        # ✅ Find the downloaded Smartsheet Excel file
        sheet_folder = os.path.abspath(f"sheets/{sheet_id}")
        mapping_folder = os.path.abspath(f"row_mapping/{sheet_id}")
        ensure_folder(mapping_folder)
        excel_files = glob.glob(os.path.join(sheet_folder, "*.xlsx"))

        if not excel_files:
            print(f"❌ Smartsheet Excel not found in {sheet_folder}")
            return None

        # ✅ Find the saved Excel file using wildcard (*.xlsx)
        # ✅ Read Comments sheet
        retry_count = 5
        excel_files = []
        while retry_count > 0:
            excel_files = glob.glob(os.path.join(sheet_folder, original_file))
            if excel_files:
                break
            print("⚠️ Waiting for Excel file to be available...")
            time.sleep(2)
            retry_count -= 1

        if not excel_files:
            print(f"❌ No Excel file found in {sheet_folder} after download.")
            return None
        
        # ✅ Load Excel into Pandas Safely
        with pd.ExcelFile(original_file, engine="openpyxl") as xls:
            df_comments = pd.read_excel(xls, sheet_name="Comments", header=None)

        if "Comments" not in xls.sheet_names:
            print(f"⚠️ No 'Comments' sheet found in {excel_path}")
            return None
        
        if df_comments.empty:
            print(f"⚠️ No comments found in 'Comments' sheet for {sheet_id}.")
            return None

        # ✅ Assign headers dynamically (Handle missing headers)
        expected_columns = ["Relative Row", "Comments", "Created By", "Created On", "Actual Row ID"]
        df_comments = df_comments.iloc[:, :len(expected_columns)]  # Trim extra columns
        df_comments.columns = expected_columns[:df_comments.shape[1]]  # Assign headers

        # ✅ Fetch Smartsheet row IDs from API
        row_mapping = fetch_smartsheet_row_ids(sheet_id)
    

        # ✅ Extract numeric row numbers from "Relative Row"
        df_comments["Relative Row"] = df_comments["Relative Row"].astype(str).str.extract(r"(\d+)").astype(float).astype("Int64")


        # ✅ Map "Relative Row" to "Actual Row ID" using Smartsheet row numbers
        df_comments["Actual Row ID"] = df_comments["Relative Row"].map(row_mapping)

        # ✅ Create a dictionary mapping "Relative Row" to "Actual Row ID"
        mapping_table = df_comments.set_index("Relative Row")["Actual Row ID"].to_dict()

        # ✅ Convert to DataFrame
        df_mapping = pd.DataFrame(mapping_table.items(), columns=["Relative Row", "Row ID"])

        # ✅ Save to file
        mapping_path = os.path.join(mapping_folder, f"{sheet_id}_relative_row_mapping.xlsx")
        df_mapping.to_excel(mapping_path, index=False)

        # ✅ Delete the original downloaded file after modification
        
        if os.path.exists(original_file):
            os.remove(original_file)
            print(f"🗑️ Deleted original Excel file: {original_file}")

        print(f"✅ Created Relative Row → Row ID mapping table: {mapping_path}")
        return df_mapping

    except Exception as e:
        print(f"❌ Error creating mapping table for Sheet {sheet_id}: {e}")
        return None



def merge_comments_with_row_mapping(sheet_id):
    """Merges the comments table with row mapping table using wildcard search."""
    try:
        # ✅ Define the folder path
        comments_folder = os.path.abspath(f"comments/{sheet_id}")
        row_mapping_folder = os.path.abspath(f"row_mapping/{sheet_id}")
        
        # ✅ Find the comments file using wildcard
        comments_files = glob.glob(os.path.join(comments_folder, f"{sheet_id}*_comments.xlsx"))
        mapping_files = glob.glob(os.path.join(row_mapping_folder, f"{sheet_id}*_relative_row_mapping.xlsx"))
        
        if not comments_files or not mapping_files:
            print(f"❌ Comments or Mapping file not found in {sheet_folder}")
            return None
        
        comments_file = comments_files[0]
        mapping_file = mapping_files[0]
        
        # ✅ Load the comments and mapping data
        df_comments = pd.read_excel(comments_file)
        df_mapping = pd.read_excel(mapping_file)
        
        # ✅ Ensure correct column names before merging
        df_comments.rename(columns={
            df_comments.columns[0]: "Relative Row",
            df_comments.columns[1]: "Comments",
            df_comments.columns[2]: "Created By",
            df_comments.columns[3]: "Created On"
        }, inplace=True)
        df_comments["Relative Row"] = df_comments["Relative Row"].astype(str).str.extract(r"(\d+)").astype(float).astype("Int64")
      
        
        df_mapping.rename(columns={
            df_mapping.columns[0]: "Relative Row",
            df_mapping.columns[1]: "Row ID"
        }, inplace=True)
        df_mapping['Relative Row'] = df_mapping['Relative Row'].astype("Int64")
        
        # ✅ Merge comments with row mapping
        df_merged = df_comments.merge(df_mapping, on="Relative Row", how="left")
        
        # ✅ Add Sheet ID column
        df_merged.insert(0, "Sheet ID", sheet_id)
        
        # ✅ Save the updated comments table
        merged_file_path = os.path.join(comments_folder, f"{sheet_id}_comments.xlsx")
        df_merged.to_excel(merged_file_path, index=False)
        
        print(f"✅ Merged comments saved: {merged_file_path}")
        return merged_file_path
    except Exception as e:
        print(f"❌ Error merging comments with row mapping for {sheet_id}: {e}")
        return None

def get_or_create_drive_folder(folder_name, parent_folder_id):
    """Checks if a folder exists in Google Drive, creates it if not, and returns its ID."""
    try:
        query = f"name='{folder_name}' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
        results = drive_service.files().list(q=query, fields="files(id)").execute()

        if results.get("files"):
            return results["files"][0]["id"]  # ✅ Return existing folder ID

        # ✅ Create folder if it doesn't exist
        file_metadata = {
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_folder_id]
        }
        folder = drive_service.files().create(body=file_metadata, fields="id").execute()
        return folder["id"]

    except Exception as e:
        print(f"❌ Error creating Google Drive folder {folder_name}: {e}")
        return None
    

def upload_to_google_drive(sheet_id):
    """Uploads an Excel file to Google Drive in sheets/{sheet_id} folder."""
    try:
        # ✅ Find the downloaded Excel file using wildcard
        sheet_folder = os.path.abspath(f"sheets/{sheet_id}")
        excel_files = glob.glob(os.path.join(sheet_folder, "*.xlsx"))
        if not excel_files:
            print(f"❌ Smartsheet Excel not found in {sheet_folder}")
            return None

        file_path = excel_files[0]  # ✅ Select first found file

        # ✅ Ensure `sheets/{sheet_id}` folder exists in Google Drive
        drive_sheet_folder_id = get_or_create_drive_folder(str(sheet_id), GOOGLE_DRIVE_SHEETS_FOLDER_ID)

        if not drive_sheet_folder_id:
            print(f"❌ Failed to create/find folder in Google Drive for Sheet {sheet_id}")
            return None

        # ✅ Upload the file to `sheets/{sheet_id}` folder in Drive
        file_metadata = {
            "name": os.path.basename(file_path),
            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "parents": [drive_sheet_folder_id],
        }
        media = MediaFileUpload(file_path, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        file = drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()

        print(f"✅ Uploaded {file_path} to Google Drive folder: sheets/{sheet_id}")
        return file.get("id")

    except Exception as e:
        print(f"❌ Error uploading {file_path} to Google Drive: {e}")
        return None

# ✅ Send Data to AppSheet
def send_data_to_appsheet_database(google_sheet_id, sheet_name):
    """Fetches data from Google Sheets and sends it to AppSheet Database."""
    try:
        # ✅ Fetch Google Sheets Data (Ensuring it remains an Excel file)
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{google_sheet_id}/values/{sheet_name}!A1:Z1000"
        headers = {"Authorization": f"Bearer {credentials.token}"}
        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            print(f"❌ Failed to fetch Google Sheet data: {response.text}")
            return

        sheet_data = response.json().get("values", [])
        if not sheet_data:
            print("⚠️ No data found in Google Sheet.")
            return

        # ✅ Format Data for AppSheet
        headers = sheet_data[0]
        rows_data = sheet_data[1:]

        records = []
        for row in rows_data:
            record = {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers))}
            records.append(record)

        # ✅ Send to AppSheet
        appsheet_url = f"https://api.appsheet.com/api/v2/apps/{APPSHEET_APP_ID}/tables/{APPSHEET_TABLE_NAME}/Action"
        payload = {"Action": "AddOrUpdate", "Properties": {"Locale": "en-US"}, "Rows": records}
        appsheet_headers = {"Content-Type": "application/json", "ApplicationAccessKey": APPSHEET_API_KEY}

        response = requests.post(appsheet_url, headers=appsheet_headers, json=payload)
        if response.status_code == 200:
            print(f"✅ Successfully synced data with AppSheet.")
        else:
            print(f"❌ Failed to sync with AppSheet: {response.text}")
    except Exception as e:
        print(f"❌ Error syncing with AppSheet: {e}")

# ✅ **Main Execution**
sheet_id = 457130802210692  # Replace with actual Smartsheet ID
excel_path,original_file = download_smartsheet_as_excel(sheet_id)
if excel_path:
    extract_and_store_comments(sheet_id,original_file)
    create_relative_row_mapping(sheet_id)
    merge_comments_with_row_mapping(sheet_id)
    drive_file_id = upload_to_google_drive(sheet_id)

