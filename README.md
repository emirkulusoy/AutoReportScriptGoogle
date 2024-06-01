# Automated Field Test Reporting Script

## Overview

This project provides a reliable zero-cost ticketing and auto-reporting system where you can store all outputs (text notes, pictures, etc.) on a Google Drive and a well-formatted table on a Google Spread Sheet. It creates automatic pdf reports, stores them in the same Google Drive directory, and publishes them to a preselected email distribution list to keep everyone posted.

Inputs are collected via a Google Form: "example_GoogleForm.pdf"
Table data is stored on Google Spread Sheet: "example_google_spread_sheet_table_where_stores_all_data.png"
Auto reports are published via email: "example_email_output.png" and "example_pdf_output_TMOB0909_pre-log-in_Alarm TS_6_1_2024 15_03_17.pdf"
All data is stored on Google Drive

For any questions or issues, please contact Emir K Ulusoy via emir.kursad.ulusoy@gmail.com
Feel free to adjust any section to better suit your needs.

## Features

- Automatically generates a report document from form submissions.
- Saves the report in a specified Google Drive folder.
- Sends the report as an email attachment.
- Organizes photos uploaded via form submissions into the report.

### Prerequisites

- A Google account.
- A Google Form to collect data.
- A Google Drive folder to store reports.
- A Google Docs template for the report.

### Configuration

1. **Template Setup:**
   - Create or upload a Google Docs template and get its file ID. You can use the attached example_template_file.docx as a reference.

2. **Script Deployment:**
   - Create a Google Form (docs.google.com/forms) like the attached example_GoogleForm.pdf. Feel free to add more details based on your requirements.
   - Go to the "response" part and attach a Google Spread Sheet. 
   - Go to this Google Spread Sheet and select Extensions, then "Apps Script".
   - Copy and paste the script into the editor.
   - Update the constants in the script with your specific values.


3. **Trigger Setup:**
   - Go to Triggers (script.google.com) and add a new trigger.
	   - Edit Trigger for this script
	   - Choose which function to run: onFormSubmit
	   - Which runs at deployment: Head
	   - Select event source: Form spreadsheet
	   - Select event type: On form submit
	   - Failure notification settings: notify me immediately
