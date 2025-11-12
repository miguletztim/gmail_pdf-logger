# Gmail PDF Logger

A Google Apps Script that automatically extracts PDF attachments from Gmail messages, saves them to Google Drive, and logs relevant data in a Google Sheet.  

---

## Features

- Search Gmail for emails with a specific label.
- Extract PDF attachments and save them to a designated Google Drive folder.
- Append metadata to a Google Sheet:
  - Iterative ID
  - File name
  - Email date
  - Subject
  - Sender
  - Drive link
- Automatically increments ID for new files.
- Removes the label from processed emails.

---

## How It Works

1. **Gmail Search**: The script searches Gmail using a label query defined in `SCRIPT_CONFIG.GMAIL_SEARCH_QUERY`.
2. **PDF Extraction**: All PDF attachments in the matched emails are saved to a Google Drive folder specified in `SCRIPT_CONFIG.DRIVE_FOLDER_ID`.
3. **Logging**: Metadata of each PDF is appended to the sheet defined in `SCRIPT_CONFIG.SPREADSHEET_ID` and `SCRIPT_CONFIG.SHEET_NAME`.
4. **Label Management**: After processing, the script removes the label from the email to avoid duplicate processing.

---

## Setup

1. **Create a Google Sheet** to log the PDF data.
2. **Create a Google Drive folder** to store the PDFs.
3. **Set up the script**:
   - Open [Google Apps Script](https://script.google.com/).
   - Copy the script into a new project.
   - Update `SCRIPT_CONFIG` with your:
     - Gmail search query (`GMAIL_SEARCH_QUERY`)
     - Drive folder ID (`DRIVE_FOLDER_ID`)
     - Spreadsheet ID (`SPREADSHEET_ID`)
     - Sheet name (`SHEET_NAME`)
     - File prefix (`FILE_PREFIX`)
4. **Run the script** and authorize permissions for Gmail, Drive, and Sheets.

---

## Configuration Example

```javascript
const SCRIPT_CONFIG = {
  TARGET_MAILBOX_EMAIL: "me",
  GMAIL_SEARCH_QUERY: "label:Anhang_speichern",
  DRIVE_FOLDER_ID: "YOUR_DRIVE_FOLDER_ID",
  SPREADSHEET_ID: "YOUR_SPREADSHEET_ID",
  SHEET_NAME: "Posteingangsbuch",
  FILE_PREFIX: "PE_",
};
