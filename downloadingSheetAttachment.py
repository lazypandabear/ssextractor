import smartsheet
import urllib.request
import os
from dotenv import load_dotenv

# ✅ Initialize Smartsheet API Client
load_dotenv()
SMARTSHEET_API_KEY = os.getenv("SMARTSHEET_API_KEY")
smart = smartsheet.Smartsheet(SMARTSHEET_API_KEY)

# ✅ Specify your Smartsheet Sheet ID
sheet_id = 457130802210692  # Replace with actual sheet ID

# ✅ Fetch all attachments from the sheet
att_list = smart.Attachments.list_all_attachments(sheet_id, include_all=True)
print(type(att_list.data))

# ✅ Set the destination folder for downloads
dest_dir = "C:\\Users\\Dennis\\Downloads\\Smartsheet_Attachments\\"  # Change as needed
os.makedirs(dest_dir, exist_ok=True)  # Ensure folder exists

for attach in att_list.data:
    att_id = attach.id  # Get the attachment ID
    att_name = attach.name  # Get the attachment name

    # ✅ Fetch the latest downloadable URL
    retrieve_att = smart.Attachments.get_attachment(sheet_id, att_id)
    dwnld_url = retrieve_att.url  # URL expires in 5-10 mins

    if dwnld_url:
        dest_file = os.path.join(dest_dir, att_name)  # ✅ Construct file path
        print(f"📥 Downloading '{att_name}' to {dest_file}...")

        try:
            urllib.request.urlretrieve(dwnld_url, dest_file)  # Download and save the file
            print(f"✅ Successfully downloaded '{att_name}' to {dest_file}")
        except Exception as e:
            print(f"❌ Failed to download '{att_name}': {e}")
    else:
        print(f"⚠️ No valid URL found for '{att_name}', skipping...")