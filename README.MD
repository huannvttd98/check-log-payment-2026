# Gmail to Google Sheets for Vietcombank Transactions

This project contains a Google Apps Script designed to automate the extraction of transaction details from Vietcombank (VCBDigibank) email notifications and save them into a Google Sheet.

## Key Features

- **Email Scanning**: Automatically scans Gmail for threads based on a configurable query (default: emails after 2025/01/01).
- **Filtering**: Specifically targets emails sent from "VCBDigibank" to ensure only relevant bank transactions are processed.
- **Duplicate Prevention**: Checks the Google Sheet for existing Message IDs to ensure no transaction is recorded twice.
- **Data Extraction**: Parses the email plain text body to extract key transaction details:
  - Message ID, From Address, Subject, Date
  - Beneficiary Name
  - Amount
  - Currency
  - Beneficiary Bank Name
  - Payment Details
- **Storage**: Appends the extracted information as a new row in a specified Google Sheet.

## Configuration

The script uses a `CONFIG` object at the top of the file for easy customization:

```javascript
const CONFIG = {
  QUERY: "after:2025/01/01", // Gmail search query
  SHEET_NAME: "GmailData",   // Name of the sheet tab
  SHEET_ID: "YOUR_SHEET_ID", // The ID of the destination Google Sheet
  MAX_THREADS: 100,          // Number of threads to process per run
  RAW_LIMIT: 5000,           // Character limit for raw content
};
```

## Setup & Usage

1. Open **Google Apps Script** (script.google.com) or the script editor attached to your Google Sheet.
2. Copy the content of `gmail_reader.gs` into the editor.
3. Update the `CONFIG.SHEET_ID` with the ID of your target Google Sheet.
4. Run the `runReadGmail` function.
5. (Optional) Set up a **Time-driven trigger** to run `runReadGmail` automatically (e.g., every hour or day).

## Output Structure

The script will automatically create the sheet if it doesn't exist and add headers. The columns populated are:
1. MessageID
2. From
3. Subject
4. Date
5. Beneficiary Name
6. Amount
7. Currency
8. Beneficiary Bank
9. Payment Detail
