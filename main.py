import os
import smartsheet
from ssextractor import (
    download_smartsheet_as_excel,
    extract_and_store_comments,
    create_relative_row_mapping,
    merge_comments_with_row_mapping,
    download_smartsheet_attachments,
    prepare_sheet_for_drive_upload,
    upload_to_google_drive,
    upload_comments_to_drive,
    upload_attachments_to_drive
)
from getSsSheetID import get_sheets_in_folder

def run_migration(config):
    """
    Updates configuration from the provided dictionary and runs the migration.
    Expected keys in config:
        - smartsheet_api_key
        - smartsheet_folder_id
        - google_drive_sheets_folder_id
        - google_drive_comments_folder_id
        - google_drive_attachments_folder_id
    """
    # Update environment variables (or pass the config to your functions if preferred)
    os.environ["SMARTSHEET_API_KEY"] = config.get("smartsheet_api_key")
    os.environ["GOOGLE_DRIVE_SHEETS_FOLDER_ID"] = config.get("google_drive_sheets_folder_id")
    os.environ["GOOGLE_DRIVE__COMMENTS_FOLDER_ID"] = config.get("google_drive_comments_folder_id")
    os.environ["GOOGLE_DRIVE_ATTACHMENTS_FOLDER_ID"] = config.get("google_drive_attachments_folder_id")
    
    # Depending on your original code, you might also set or use the Smartsheet folder ID:
    smartsheet_folder_id = config.get("smartsheet_folder_id")
    
    # Get sheets in the specified Smartsheet folder
    sheets, sheet_info, sheet_ids_list = get_sheets_in_folder(smartsheet_folder_id)
    print(f"🔄 Found {len(sheets)} sheets in folder ID {smartsheet_folder_id}.")
    for sheet in sheets:
        sheet_id = sheet.id
        # Execute the various migration functions
        download_smartsheet_as_excel(sheet_id)
        extract_and_store_comments(sheet_id)
        create_relative_row_mapping(sheet_id)
        merge_comments_with_row_mapping(sheet_id)
        download_smartsheet_attachments(sheet_id)
        prepare_sheet_for_drive_upload(sheet_id)
        upload_to_google_drive(sheet_id)
        upload_comments_to_drive(sheet_id)
        upload_attachments_to_drive(sheet_id)
    
    print("🎉 Migration Completed Successfully!")
    return "Migration Completed Successfully!"

# If you want to allow running the script directly as well:
if __name__ == '__main__':
    # You might define default values or load from environment here.
    config = {
        "smartsheet_api_key": os.getenv("SMARTSHEET_API_KEY"),
        "smartsheet_folder_id": "YOUR_DEFAULT_FOLDER_ID",
        "google_drive_sheets_folder_id": os.getenv("GOOGLE_DRIVE_SHEETS_FOLDER_ID"),
        "google_drive_comments_folder_id": os.getenv("GOOGLE_DRIVE__COMMENTS_FOLDER_ID"),
        "google_drive_attachments_folder_id": os.getenv("GOOGLE_DRIVE_ATTACHMENTS_FOLDER_ID")
    }
    run_migration(config)
