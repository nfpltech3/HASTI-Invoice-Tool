# HASTI Invoice to CSV User Guide

## Introduction

This application automates the conversion of **HASTI Petro Chemical DO (Delivery Order) invoices** into the specific **Logisys CSV format**. It parses PDF files, extracts key financial data, and maps Bill of Entry numbers to Job Numbers using a Job Register.

**Who is this for?** Accounts team members at Nagarkot Forwarders who need to upload HASTI purchase entries to the Logisys system.

**Key Features:**
- **PDF Extraction**: Reads Vendor Invoice No, Date, BOE No, BL No, and amounts directly from PDF files.
- **Intelligent Classification**: Automatically detects "Transport" vs. "CFS" invoices to apply correct tax codes and charge names.
- **Job Number Mapping**: Cross-references BOE numbers with your Job Register.
- **Logisys Compliance**: Outputs a 41-column CSV ready for direct upload.

---

## How to Use

### 1. Launching the App
1. Locate the application folder.
2. Double-click `HASTI_Invoice_to_CSV.exe`.
3. The app will launch in full-screen mode with the HASTI branding header.

### 2. The Workflow (Step-by-Step)

1. **Select Invoice PDFs**: 
   - Click the **"Select Invoice PDFs"** button.
   - Select one or more `.pdf` invoice files. 
   - *Tip: You can select multiple files at once by holding Ctrl or Shift.*

2. **Select Job Register**:
   - Click the **"Select Job Register"** button.
   - Choose your Job Register file (`.csv`, `.xlsx`, or `.xls`).
   - *Requirement:* The file must have columns for **"BE No"** and **"Job No"**.

3. **Process & Generate CSV**:
   - Click the blue **"▶ Process & Generate CSV"** button.
   - The application will process each PDF, calculate taxes, and generate the output.
   - Watch the **Processing Log** for real-time status updates.

4. **Output**:
   - The CSV file is saved automatically in the `HASTI_Output` folder.
   - Filename format: `Hasti_YYYY-MM-DD_HH-MM.csv`.

---

## Logic & Business Rules

### Invoice Type Detection
The app scans the PDF text for the phrase **"TRANSPORTATION OF GOODS - ROAD"**.

| Invoice Type | Detection Rule | Charge Name | Tax Type | Avail Tax Credit |
| :--- | :--- | :--- | :--- | :--- |
| **Transport** | "TRANSPORTATION..." found | `Transport Charges _ FCM (E)` | Taxable | 100 |
| **CFS** | Phrase NOT found | `CFS CHARGES (1)` | Pure Agent | No |

### Tax Calculations
- **Transport**: CGST (6%), SGST (6%), WH Tax (2% of Taxable).
- **CFS**: No GST applied (Pure Agent). WH Tax (2% of Taxable).

---

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Select Invoice PDFs** | File picker for source invoice documents. | `.pdf` |
| **Select Job Register** | File picker for BOE-to-Job mapping. | `.csv`, `.xlsx` |
| **▶ Process** | Triggers the conversion logic. | Button Action |
| **Processing Log** | Displays extraction details and errors. | Text Output |

---

## Troubleshooting

| Message / Issue | Possible Cause | Solution |
| :--- | :--- | :--- |
| **"No PDFs selected"** | You clicked Process without choosing invoices. | Click "Select Invoice PDFs" first. |
| **"No job register selected..."** | You clicked Process without a Job Register. | Click "Select Job Register" first. |
| **"No match found" (Ref No)** | The BOE No in the invoice isn't in the Job Register. | Update your Job Register with the missing BOE. |
| **"Failed to extract text..."** | The PDF might be a scanned image or corrupted. | Ensure the PDF is text-readable (not just an image). |
| **"No valid data extracted..."** | The PDF layout doesn't match the expected HASTI invoice format. | Check if the invoice format has changed. |
