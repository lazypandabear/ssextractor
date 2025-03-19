import smartsheet
import os
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from dotenv import load_dotenv

# Load API Keys
load_dotenv()
SMARTSHEET_API_KEY = os.getenv("SMARTSHEET_API_KEY")

# Initialize Smartsheet Client
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_API_KEY)

# Get all sheets
response = smartsheet_client.Sheets.list_sheets(include_all=True)
sheets = response.data

for sheet in sheets:
    sheet_id = sheet.id
    sheet_name = sheet.name

    # Get Sheet Data
    sheet_data = smartsheet_client.Sheets.get_sheet(sheet_id)
    columns = [col.title for col in sheet_data.columns]
    
    # Extract row data
    rows = []
    for row in sheet_data.rows:
        row_data = {columns[i]: cell.value if cell.value else '' for i, cell in enumerate(row.cells)}
        rows.append(row_data)

    # Convert to DataFrame
    df = pd.DataFrame(rows)
    df.to_csv(f"{sheet_name}.csv", index=False)

    print(f"✅ Extracted data for: {sheet_name}")
