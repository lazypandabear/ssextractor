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
import urllib.parse
from ssextractor import *
from getSsSheetID import get_sheets_in_folder

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


def cleanup_downloads(base_dir):
    """Deletes all extracted files from the local PC after successful upload."""
    try:
        if os.path.exists(base_dir):
            shutil.rmtree(base_dir)  # Remove the entire folder and its contents
            print(f"🗑️ Successfully deleted extracted files from {base_dir}")
        else:
            print(f"⚠️ Cleanup skipped: {base_dir} does not exist.")
    except Exception as e:
        print(f"❌ Error while deleting files: {e}")

# ✅ Get All Sheets from Smartsheet
#response = smartsheet_client.Sheets.list_sheets(include_all=True)
#sheets = response.data

# ✅ Get All Sheets from a Smartsheet Folder
folder_id = 5778260903651204
sheets,sheet_info,sheet_ids_list =get_sheets_in_folder(folder_id)


for sheet in sheets:
    sheet_id = sheet.id # Replace with actual Smartsheet ID
    sheet_name = sheet.name
    print(f"🔄 Processing: {sheet_name} (ID: {sheet_id})")
# ✅ **Main Execution**
    download_smartsheet_as_excel(sheet_id)
    extract_and_store_comments(sheet_id)
    create_relative_row_mapping(sheet_id)
    merge_comments_with_row_mapping(sheet_id)
    download_smartsheet_attachments(sheet_id)
    prepare_sheet_for_drive_upload(sheet_id)
    upload_to_google_drive(sheet_id)
    upload_comments_to_drive(sheet_id)
    upload_attachments_to_drive(sheet_id)

cleanup_downloads("C:/users/Dennis/VSCode/smartsheetAPI/attachments")  
cleanup_downloads("C:/users/Dennis/VSCode/smartsheetAPI/comments")  
cleanup_downloads("C:/users/Dennis/VSCode/smartsheetAPI/row_mapping")  
cleanup_downloads("C:/users/Dennis/VSCode/smartsheetAPI/sheets")  
print("🎉 Migration Completed Successfully!")