import smartsheet
import os
import pandas as pd
from dotenv import load_dotenv

# ✅ Load environment variables
load_dotenv(override=True)
SMARTSHEET_API_KEY = os.getenv("SMARTSHEET_API_KEY")

# ✅ Initialize Smartsheet Client
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_API_KEY)

def get_sheets_in_folder(folder_id):
    """Retrieves all sheets inside a given Smartsheet folder and returns them as a list of dictionaries."""
    try:
        # ✅ Get the folder details
        folder = smartsheet_client.Folders.get_folder(folder_id)
        sheets = folder.sheets

        if not sheets:
            print(f"⚠️ No sheets found in Folder ID {folder_id}.")
            return []

        sheet_info = [{"Sheet ID": sheet.id, "Sheet Name": sheet.name} for sheet in sheets]
        sheet_ids_list = [sheet.id for sheet in sheets]

        print(f"✅ Found {len(sheets)} sheets in Folder ID {folder_id}:")
        for sheet in sheet_info:
            print(f"  - {sheet['Sheet Name']} (ID: {sheet['Sheet ID']})")

        return sheets,sheet_info,sheet_ids_list

    except smartsheet.exceptions.ApiError as e:
        print(f"❌ Smartsheet API error: {e}")
        return None
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return None

def save_sheet_ids_to_csv(folder_id, output_folder="sheet_id_exports"):
    """Extracts all sheet IDs from a Smartsheet folder and saves them as a CSV file."""
    try:
        os.makedirs(output_folder, exist_ok=True)  # ✅ Ensure output folder exists
        sheets = get_sheets_in_folder(folder_id)

        if sheets:
            # ✅ Convert to DataFrame
            df = pd.DataFrame(sheets)

            # ✅ Define output file path
            csv_filename = f"{output_folder}/sheet_ids_{folder_id}.csv"
            df.to_csv(csv_filename, index=False, encoding="utf-8")

            print(f"✅ Saved all Sheet IDs from Folder {folder_id} to {csv_filename}")
            return csv_filename
        else:
            print(f"⚠️ No sheets found, skipping CSV creation.")
            return None

    except Exception as e:
        print(f"❌ Error saving Sheet IDs to CSV: {e}")
        return None

if __name__ == "__main__":
    folder_id = 5778260903651204  # Replace with the actual folder ID
    print(get_sheets_in_folder(folder_id))
    #save_sheet_ids_to_csv(folder_id)