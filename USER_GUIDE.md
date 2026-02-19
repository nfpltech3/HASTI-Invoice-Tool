# HASTI Tools User Guide

## Introduction

This package contains **two desktop applications** for converting HASTI Petro Chemical financial documents into Logisys-compatible CSV files:

1. **HASTI DO Invoice to CSV Converter** — Parses DO (Delivery Order) invoice PDFs and generates a purchase CSV.
2. **Ledger to Purchase CSV Converter** — Converts Excel ledger reports into a purchase CSV, cross-referencing a Job Register for BOE-to-Job mapping.

**Who is this for?** Accounts team members at Nagarkot Forwarders who need to upload purchase entries to the Logisys system.

**Key Features:**
- Automatic extraction of invoice fields (Invoice No, Date, BOE No, amounts, taxes) from PDF invoices
- Automatic detection of Transport vs. CFS invoice types with different tax treatments
- Job Register lookup to map BOE numbers to Job Numbers
- Output in the exact 41-column Logisys CSV format
- Branded Nagarkot GUI with real-time processing logs

---

## How to Use

### 1. Launching the Apps

- Locate the application folder.
- Double-click `HASTI_Invoice_to_CSV.exe` or `Ledger_to_CSV.exe`.
- The app will launch in full-screen mode with the Nagarkot branding header.

---

## Tool 1: HASTI DO Invoice to CSV Converter

### 2. The Workflow (Step-by-Step)

1. **Select Invoice PDFs**: Click the **"Select Invoice PDFs"** button → Navigate and select one or more HASTI DO invoice PDF files. You can select multiple files at once.
   - *Supported Format:* `.pdf`
   - *Note: PDFs must be HASTI Petro Chemical invoices with standard fields (Invoice No, BOE No, Total Amount, etc.)*

2. **Select Job Register**: Click the **"Select Job Register"** button → Select your Job Register file.
   - *Supported Formats:* `.csv`, `.xlsx`, `.xls`
   - *Note: The file must contain columns named `BE No` and `Job No` (case-insensitive, spaces trimmed).*

3. **Process & Generate CSV**: Click the blue **"▶ Process & Generate CSV"** button.
   - The app will parse each PDF, extract invoice details, match BOE numbers to Job Numbers, and generate the CSV.
   - Progress and any issues are displayed in the **Processing Log** panel.

4. **Output**: The CSV is automatically saved to `HASTI_Output/` folder in the same directory as the app.
   - Filename format: `Hasti_YYYY-MM-DD_HH-MM.csv`
   - A success popup confirms the save location and record count.

### How Invoice Type Detection Works

The app automatically detects the invoice type by searching for the phrase **"TRANSPORTATION OF GOODS - ROAD"** in the PDF text:

| Invoice Type | Charge Name | Tax Type | SAC Code | Avail Tax Credit |
| :--- | :--- | :--- | :--- | :--- |
| **Transport** (phrase found) | Transport Charges _ FCM (E) | Taxable | 996793 | 100 |
| **CFS** (phrase not found) | CFS CHARGES (1) | Pure Agent | 996711 | No |

- **Transport invoices**: CGST and SGST are calculated at 6% each of Total Amount. WH Tax at 2%.
- **CFS invoices**: No tax codes applied. WH Tax at 2% of Total Amount.

---

## Tool 2: Ledger to Purchase CSV Converter

### 2. The Workflow (Step-by-Step)

1. **Select Job Register**: Click **"Select Job Register"** → Select your Job Register file (`.csv` or `.xlsx`).
   - *Note: Must be selected FIRST before the Ledger Report.*
   - *Required columns:* Any of `BOE No`, `BE No.`, `BE No`, `BOE No.`, `BOE Number`, `Bill of Entry No` AND any of `Job No.`, `Job No`, `Job Number`, `Ref No`, `Reference No`.

2. **Select Ledger Report**: Click **"Select Ledger Report"** → Select the ledger dump Excel file.
   - *Supported Format:* `.xlsx` only
   - *Required columns:* `Receipt No.`, `BOE No.`, `Txn Date`, `Consignee Name`

3. **Process & Generate CSV**: Click the blue **"▶ Process & Generate CSV"** button.
   - Rows missing `Receipt No.` or `BOE No.` are skipped and logged.
   - Each row is matched against the Job Register for the Job Number.

4. **Output**: CSV saved to `Kale Output/` folder.
   - Filename format: `purchase_DD-MM-YY HH-MM.csv`

### Consignee-Based Charge Logic

| Consignee | Charge Name | Amount | Tax | Avail Tax Credit |
| :--- | :--- | :--- | :--- | :--- |
| Starts with **"ABBOTT HEALTHCARE"** | GATE PASS CHARGES - REIM | 336 | None | No |
| **All others** | GATE PASS CHARGES CCL | 285 | CGST ₹25.65 + SGST ₹25.65 | 100 |

---

## Interface Reference

### HASTI Invoice to CSV

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Select Invoice PDFs** | Opens file picker for one or more invoice PDFs | `.pdf` files |
| **Select Job Register** | Opens file picker for the BOE-to-Job mapping file | `.csv`, `.xlsx`, `.xls` |
| **▶ Process & Generate CSV** | Triggers parsing, matching, and CSV generation | Button action |
| **Processing Log** | Real-time log of extraction, matching, and errors | Read-only text |
| **Status Label** | Shows "Ready", "Processing...", "Completed Successfully", or error | Text indicator |
| **Exit** | Closes the application | Button (footer) |

### Ledger to CSV

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Select Job Register** | Opens file picker for BOE-to-Job mapping file | `.csv`, `.xlsx` |
| **Select Ledger Report** | Opens file picker for the ledger Excel dump | `.xlsx` |
| **▶ Process & Generate CSV** | Triggers conversion and CSV generation | Button action |
| **Processing Log** | Real-time log showing row-by-row processing | Read-only text |
| **Status Label** | Current processing state | Text indicator |
| **Exit** | Closes the application | Button (footer) |

---

## Output Locations

| Tool | Output Folder | Filename Pattern |
| :--- | :--- | :--- |
| HASTI Invoice to CSV | `HASTI_Output/` | `Hasti_2026-02-18_17-00.csv` |
| Ledger to CSV | `Kale Output/` | `purchase_18-02-26 17-00.csv` |

Both folders are created automatically in the same directory as the `.exe` file.

---

## Troubleshooting & Validations

If you see an error, check this table:

### HASTI Invoice to CSV

| Message | What it means | Solution |
| :--- | :--- | :--- |
| **"No PDFs selected"** | Process was clicked without selecting any PDF files | Click "Select Invoice PDFs" first and choose at least one `.pdf` file |
| **"No job register selected or loaded"** | Job Register file was not selected or could not be parsed | Click "Select Job Register" and ensure file has `BE No` and `Job No` columns |
| **"Failed to extract text from [filename]"** | pdfplumber could not read the PDF (corrupted or image-only) | Re-download the PDF or ensure it is text-based, not a scanned image |
| **"No valid data extracted from [filename]"** | Regex patterns did not match any fields in the PDF text | Verify the PDF is a HASTI Petro Chemical DO invoice in the expected format |
| **"No valid data extracted from PDFs"** | None of the selected PDFs yielded usable data | Check all PDFs are valid HASTI invoices |
| **"Failed to write CSV: [error]"** | File system error during CSV save | Ensure `HASTI_Output` folder isn't open/locked and disk has space |
| **"Regex extraction error: [error]"** | An unexpected field format was encountered | Check the Processing Log for the specific field that failed |
| **"No match found"** in Ref No column | BOE number from invoice was not found in the Job Register | Verify the Job Register contains the BOE number for this invoice |

### Ledger to CSV

| Message | What it means | Solution |
| :--- | :--- | :--- |
| **"Please select Job Register file before selecting Ledger Report"** | Attempted to pick Ledger before Job Register | Always select Job Register first |
| **"BOE column not found in Job Register file"** | None of the expected column names found | Ensure file has a column named `BOE No`, `BE No.`, `BE No`, etc. |
| **"Job No column not found in Job Register file"** | None of the expected column names found | Ensure file has a column named `Job No.`, `Job No`, `Job Number`, etc. |
| **"Skipping row [n] due to missing Receipt No"** | A ledger row has no Receipt Number | Check source Excel for incomplete data in that row |
| **"Skipping row [n] ... due to missing BOE No"** | A ledger row has no Bill of Entry Number | Ensure BOE data is populated in the source file |
| **"Skipping row [n] ... due to invalid Txn Date"** | The date in `Txn Date` column could not be parsed | Fix the date format in the source Excel file |
| **"No Job No found for BOE No: [number]"** | BOE number exists in ledger but not in Job Register | Update Job Register to include this BOE entry |
| **"No valid rows to process"** | All rows were skipped due to missing data | Review source Excel — every row needs Receipt No., BOE No., and Txn Date |
| **"Failed to create CSV: [error]"** | File system error during save | Ensure `Kale Output` folder isn't locked and disk has space |
