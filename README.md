# Smartsheet Data Extractor and Integrator

This project is designed to extract data from Smartsheet, including sheet data, comments, and attachments, and then integrate it with Google Drive and AppSheet. It automates the process of backing up Smartsheet information and making it accessible in other platforms.

## What This Project Does

Imagine you have a lot of important information stored in Smartsheet. This project will help you:

1.  **Download Smartsheet Data:** Grab the main data from your Smartsheet (like a big table) and save it as an Excel file.
2.  **Get Comments:** Extract all the comments made on your Smartsheet, who made them, and when.
3.  **Grab Attachments:** Download any files (like PDFs, images, etc.) attached to your Smartsheet.
4.  **Put it on Google Drive:** Upload all the data, comments, and attachments to your Google Drive. This makes it easy to share and access.
5.  **Send to AppSheet:** If you use AppSheet, you can automatically send your Smartsheet data there to build apps.
6.  **Get all the sheets from a folder:** if you want to extract more then 1 sheet, this project can help you.

## Before You Begin (What You Need)

Think of these as the tools you need for this project:

1.  **Smartsheet Account:** You need a Smartsheet account with a sheet you want to extract data from.
2.  **Google Account:** You need a Google account to use Google Drive.
3.  **AppSheet Account (Optional):** Only if you want to send data to AppSheet.
4.  **Python:** This code is written in Python. If you don't have it, download and install it from [https://www.python.org/downloads/](https://www.python.org/downloads/).
5.  **Code Editor:** A program to read and edit code, like VS Code, Sublime Text, or even Notepad++.
6.  **Service account:** you need to create a service account in google cloud platform.

## Setting Up (Step-by-Step)

This is how you get your tools ready:

### 1. Get Your API Keys

   *   **Smartsheet API Key:**
        1.  Log in to Smartsheet.
        2.  Click on "Account" (usually your profile picture).
        3.  Go to "Apps & Integrations."
        4.  Click on "API Access" and follow the instructions to generate an API key.
        5.  Copy and save this key somewhere safe.
   *   **Google Drive API Credentials:**
        1.  Go to the Google Cloud Console: [https://console.cloud.google.com/](https://console.cloud.google.com/)
        2.  Create a new project (or use an existing one).
        3.  Go to "APIs & Services" -> "Library."
        4.  Search for "Google Drive API" and enable it.
        5.  Search for "Google Sheets API" and enable it.
        6.  Go to "APIs & Services" -> "credentials."
        7.  create credentials -> service account.
        8.  in service account details fill the information and in step 2 grant this service account access to the project.
        9.  Step 3 is optional, click done.
        10. select the service account you just created and click "keys" then "add key" -> "create new key" -> "JSON".
        11. a file call something like "service_account.json" will download, this is your service account json.
   *   **AppSheet API Key (Optional):**
        1.  Log in to AppSheet.
        2.  Go to "My Account" -> "Integrations."
        3.  Generate an API key.
        4.  Copy and save this key.
   *   **Appsheet APP ID & TABLE NAME (Optional):**
        1. Go to your app
        2. go to manage -> integration
        3. copy the app id
        4. Copy the table name where you will send your data.

### 2. Set Up the `.env` File

   1.  In the same folder as your code files, create a new file called `.env` (no file extension).
   2.  Paste the following text into `.env`, and replace the example values with your actual API keys and IDs.
        ```properties
        SMARTSHEET_API_KEY='YOUR_SMARTSHEET_API_KEY'
        GOOGLE_DRIVE_PARENT_FOLDER_ID='YOUR_GOOGLE_DRIVE_PARENT_FOLDER_ID' #if you dont want to create folders, just put here the parent folder
        GOOGLE_DRIVE_SHEETS_FOLDER_ID='YOUR_GOOGLE_DRIVE_SHEETS_FOLDER_ID'
        GOOGLE_DRIVE__COMMENTS_FOLDER_ID='YOUR_GOOGLE_DRIVE_COMMENTS_FOLDER_ID'
        GOOGLE_DRIVE_ATTACHMENTS_FOLDER_ID='YOUR_GOOGLE_DRIVE_ATTACHMENTS_FOLDER_ID'
        SERVICE_ACCOUNT_FILE='service_account.json' #rename your key to service_account.json
        APPSHEET_API_KEY='YOUR_APPSHEET_API_KEY'
        APPSHEET_APP_ID='YOUR_APPSHEET_APP_ID'
        APPSHEET_TABLE_NAME='YOUR_APPSHEET_TABLE_NAME'
        ```
   3. Save the file.

### 3. Install Python Packages

   1.  Open a terminal or command prompt.
   2.  Navigate to the folder where your code files are located (using `cd your_folder_path`).
   3.  **Install the required packages from `requirements.txt`:** This project uses several external Python libraries. Instead of installing them one by one, you can install them all at once using the `requirements.txt` file. Run the following command:
        ```bash
        pip install -r requirements.txt
        ```
   4. if you need to install pip go here: [https://pip.pypa.io/en/stable/installation/](https://pip.pypa.io/en/stable/installation/)

### 4. Put the `service_account.json` file in the right folder:

   1.  Put the `service_account.json` key in the same folder where you have the code.

## Running the Code (Finally!)

1.  **Choose the Main Code:** You'll mostly be working with `main.py`.
2.  **sheet Id:** if you are working with only 1 sheet change `sheet_id = 457130802210692` with your sheet id in `ssextractor.py`.
3.  **Folder ID:** if you are working with more then 1 sheet you can add the folder ID in the `getSsSheetID.py` file in the `folder_id = 5778260903651204` line.
4.  **Open Terminal:** Open a terminal or command prompt in the folder where your code is.
5.  **Run:** Type the following command and press Enter:
    ```bash
    python main.py
    ```
6.  **Wait:** The code will run, and you'll see messages in the terminal telling you what it's doing.
7.  **Check Google Drive:** Once it's done, go to your Google Drive. You should see new folders and files containing your Smartsheet data, comments, and attachments.
8.  **Check AppSheet (If Used):** If you set up AppSheet, your data should be there too.

## Code Files Explained (What Does Each File Do?)

*   **`main.py`:**
    *   This is the main program that you run.
    *   It calls all the other functions in the other files.
    *   It connects to Smartsheet, gets the data, comments, and attachments.
    *   It uploads everything to Google Drive and sends data to AppSheet (if you configured it).
*   **`ssextractor.py`:**
    *   This file has all the functions for working with Smartsheet and Google Drive.
    *   It downloads Smartsheet as Excel.
    *   It extracts comments.
    *   It downloads attachments.
    *   It uploads to Google Drive.
    *   it has the function to send data to appsheet.
*   **`getSsSheetID.py`:**
    *   This file connects to Smartsheet and gets the IDs of all the sheets in a folder.
    *   it saves the id's in a list and in a .csv file if you want.
*   **`backup_ssextractor.py`:**
    *   Older version of ssextractor, it can still be useful, but is not called by the `main.py` file.
*   **`.env`:**
    *   This file stores your secret API keys and IDs.
    *   **Never share this file with anyone!**
*   **`comment_extractor.py`:**
    *   This is a standalone file that just extracts comments from Smartsheet and saves them to an Excel file.
    *   It's not used directly by `main.py`.
*   **`downloadingSheetAttachment.py`:**
    *   This is a standalone file that just downloads attachments from Smartsheet.
    *   It's not used directly by `main.py`.
*   **`getSmartsheetAsExcel.py`**
    *   This is a standalone file that just extract the sheet in excel.

## Troubleshooting (Help! It's Not Working)

*   **Check Your API Keys:** Double-check that you entered your API keys correctly in the `.env` file. Even a tiny mistake will cause problems.
*   **Folder IDs:** Make sure you're using the right Google Drive folder IDs.
*   **Correct sheet id:** verify that you are using the right sheet id in the code.
*   **correct Folder ID:** verify that you are using the right folder id in the code, it can be in `getSsSheetID.py` or in `.env` file.
*   **Permissions:** If you're having trouble with Google Drive, make sure the Google Drive API is enabled and you set up the credentials properly.
*   **service account:** verify that the service account has the right permissions.
*   **Error Messages:** Read the error messages in the terminal carefully. They often tell you exactly what's wrong.
*   **Packages:** If it says that a package is missing, make sure you installed all the packages from `requirements.txt`.
*   **File Paths:** If you're having trouble with files not being found, make sure the file paths in your code are correct.
*   **Restart:** Sometimes, just restarting your computer or terminal can fix weird issues.

## That's it!

This project is a powerful way to manage and back up your Smartsheet data. If you have any more questions, feel free to ask!

Engr. Dennis A. Garcia
Head - nexBT Solutions
