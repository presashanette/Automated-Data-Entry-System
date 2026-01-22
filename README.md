# Automated-Data-Entry-System (Python / Selenium)

### Project Overview
- This project is a robust automation tool designed to migrate patient records and visit logs from a legacy Excel database into a modern web-based Electronic Health Record (EHR) system.
- It solves the problem of manual data entry errors and significantly reduces the man-hours required to digitize physical records.

### Key Features
- Intelligent Duplicate Detection: Scans existing web records for date and diagnosis matches to prevent duplicate entries before logging new ones.
- Dynamic Form Handling: Automatically detects and handles dynamic input fields that shift position based on the number of diagnoses.
- Cross-Platform Compatibility: Logic included to handle keyboard shortcuts for both Windows (Ctrl) and MacOS (Command).
- Persistent Session Management: Uses a local Chrome profile to maintain login sessions, avoiding repetitive 2FA or login screens.
- Data Integrity: Writes status updates ("True") back to the source Excel file immediately upon success to prevent data loss in case of interruption.

### Technology Stack
- Python 3.10+
- Selenium Webdriver: For browser interaction and DOM manipulation.
- Pandas: For efficient Excel data processing and filtering.
- OpenPyXL: For precise cell-writing back to the source file without breaking formatting.

### How It Works
- Data Ingestion: Reads a "Master Patient List" and a "Transaction Log" (Visits) from Excel.
- Search & Verify: Searches the web portal for the patient.
- If found: Navigates to the profile.
- If missing: Automatically fills the registration form and creates the profile.
- Visit Logging: Iterates through patient history.
- Checks if the visit date already exists.
- If the visit exists but the diagnosis is missing, it enters Edit Mode.
- If the visit is new, it opens the New Visit modal.
- Reporting: Updates the Excel sheet in real-time as rows are completed.

### Security Note
- This repository contains the source code logic only.
- All patient data, credentials, and specific target URLs have been sanitized for privacy compliance.
