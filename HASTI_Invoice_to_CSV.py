import csv
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pdfplumber
import pandas as pd
import os
from datetime import datetime
import logging
import sys

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# Setup logging to file
logging.basicConfig(
    filename='do_invoice_processor.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger()

# --- Nagarkot Brand Color Palette ---
BG_COLOR = "#F4F6F8"
CARD_BG = "#FFFFFF"
ACCENT = "#1F3F6E"
ACCENT_HOVER = "#2A528F"
TEXT_PRIMARY = "#1E1E1E"
TEXT_SECONDARY = "#6B7280"
BORDER_COLOR = "#E5E7EB"
SUCCESS_COLOR = "#1F3F6E"
ERROR_RED = "#D8232A"
LOG_BG = "#FAFBFC"
LOG_FG = "#1E1E1E"

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Custom handler to display logs in GUI
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        try:
            if self.text_widget.winfo_exists():
                msg = self.format(record)
                self.text_widget.config(state='normal')
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.config(state='disabled')
                self.text_widget.see(tk.END)
        except Exception:
            pass

# Function to convert date formats to DD/MM/YYYY or DD-MMM-YYYY
def convert_date_format(date_str, out_fmt="%d/%m/%Y"):
    for fmt in [
        "%b %d, %Y",  # Jun 13, 2025
        "%d/%m/%Y",   # 13/06/2025
        "%d/%b/%Y",   # 13/Jun/2025
        "%d-%m-%Y",   # 13-06-2025
        "%d.%m.%Y",   # 13.06.2025
        "%d-%b-%y",   # 13-Jun-25
        "%d-%b-%Y"    # 13-Jun-2025
    ]:
        try:
            date_obj = datetime.strptime(date_str, fmt)
            return date_obj.strftime(out_fmt)
        except ValueError:
            continue
    return "Not Found"

# Function to clean numeric strings
def clean_numeric_string(value):
    if isinstance(value, (str, float, int)):
        value = str(value).replace(",", "").strip()
        try:
            float_val = float(value)
            if float_val.is_integer():
                return str(int(float_val))
            return str(float_val)
        except ValueError:
            return str(value)
    return str(value)

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path, log_callback):
    log_callback(f"Extracting text from {os.path.basename(pdf_path)}...")
    try:
        text = ""
        tables_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                if page.extract_text():
                    text += page.extract_text() + "\n"
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            row_text = [str(cell) if cell else "" for cell in row]
                            tables_data.append(row_text)
        combined_text = text + "\n" + "\n".join([" ".join(row) for row in tables_data])
        log_callback(f"Raw extracted text (first 1000 chars): {combined_text[:1000]}")
        return combined_text, tables_data
    except Exception as e:
        log_callback(f"Error extracting text from {os.path.basename(pdf_path)}: {str(e)}")
        logger.error(f"Failed to extract text from {pdf_path}: {e}")
        return None, []

def extract_invoice_details_with_regex(text, tables_data, log_callback):
    log_callback("Extracting fields for HASTI...")
    results = []
    try:
        # Extract Invoice No
        inv_no = re.search(r'Invoice No\.?\s*([A-Z0-9\/\-]+)', text)
        # Extract Invoice Date
        inv_date = re.search(r'Invoice Date\s*([0-9\-]+)', text)
        # Extract BOE No and BOE Date
        boe_match = re.search(r'BOE No\.?\s*([0-9]+)-([0-9\-]+)', text)
        if boe_match:
            boe_no = boe_match.group(1)
            boe_date = boe_match.group(2)
        else:
            boe_no = "Not Found"
            boe_date = "Not Found"
        # Extract BL No
        bl_no = re.search(r'BL No\.?\s*([A-Z0-9]+)', text)
        # Charge or GL Amount (Total Amount)
        total_amt = re.search(r'Total Amount\s*([0-9,.]+)', text)
        # CGST/SGST (take from first service line or from summary)
        cgst = re.search(r'CGST\s*[0-9%]*\s*([0-9,.]+)', text)
        sgst = re.search(r'SGST\s*[0-9%]*\s*([0-9,.]+)', text)
        # Total Invoice Amount
        total_invoice_amt = re.search(r'Total Invoice Amount\s*([0-9,.]+)', text)
        # Detect if TRANSPORTATION OF GOODS - ROAD is present (allowing for up to 100 chars between)
        match = re.search(r'TRANSPORTATION\s*OF(.{0,100}?)GOODS\s*-\s*ROAD', text, re.IGNORECASE | re.DOTALL)
        is_transport = bool(match)
        log_callback(f"DEBUG: is_transport={is_transport} (regex match for scattered 'TRANSPORTATION OF ... GOODS - ROAD' with up to 100 chars in between)")
        if match:
            log_callback(f"DEBUG: Matched text: {match.group(0)}")
        else:
            idx = text.upper().find('TRANSPORTATION')
            if idx != -1:
                snippet = text[max(0, idx-200):idx+200]
            else:
                snippet = text[:400]
            log_callback(f"DEBUG: No regex match. Large text snippet: {snippet}")

        # Always extract both Amount and WH Tax Taxable
        total_invoice_amt_val = total_invoice_amt.group(1).replace(',', '') if total_invoice_amt else "0"
        total_amt_val = total_amt.group(1).replace(',', '') if total_amt else "0"

        details = {
            "Organization": "HASTI PETRO CHEMICAL & SHIPPING LTD.",
            "Vendor Inv No": inv_no.group(1) if inv_no else "Not Found",
            "Vendor Inv Date": inv_date.group(1) if inv_date else "Not Found",
            "BOE No": boe_no,
            "BOE Date": boe_date,
            "BL No": bl_no.group(1) if bl_no else "Not Found",
            "Charge or GL Amount": total_invoice_amt_val,  # Default to Total Invoice Amount
            "Amount": total_invoice_amt_val,
            "WH Tax Taxable": total_amt_val,
            "Total Amount": total_amt_val,
            "Total Invoice Amount": total_invoice_amt_val,
            "CGST": cgst.group(1).replace(',', '') if cgst else "0",
            "SGST": sgst.group(1).replace(',', '') if sgst else "0",
            "Ref No": boe_no,
            "is_transport": is_transport
        }
        log_callback(f"Extracted HASTI details: {details}")
        results.append(details)
    except Exception as e:
        log_callback(f"Regex extraction error: {e}")
    return results

def create_csv(all_details, output_path, log_callback):
    fixed_fields = {
        "Organization Branch": "AHMEDABAD",
        "Currency": "INR",
        "ExchRate": "1",
        "Due Date": "",
        "Charge or GL": "Charge",
        "Charge or GL Name": "",
        "DR or CR": "DR",
        "Cost Center": "",
        "Branch": "GUJARAT",
        " Charge Narration": "",
        "TaxGroup": "",
        "Tax Type": "",
        "SAC or HSN": "996793",
        "Taxcode1": "",
        "Taxcode2": "",
        "Taxcode3": "",
        "Taxcode4": "",
        "Taxcode1 Amt": "",
        "Taxcode2 Amt": "",
        "Taxcode3 Amt": "",
        "Taxcode4 Amt": "",
        "Avail Tax Credit": "",
        "LOB": "CCL IMP",
        "Ref Type": "",
        "Start Date": "",
        "End Date": "",
        "WH Tax Code": "194C",
        "WH Tax Percentage": "2",
        "Round Off": "Yes",
        "CC Code": ""
    }
    column_order = [
        "Entry Date",
        "Posting Date",
        "Organization",
        "Organization Branch",
        "Vendor Inv No",
        "Vendor Inv Date",
        "Currency",
        "ExchRate",
        "Narration",
        "Due Date",
        "Charge or GL",
        "Charge or GL Name",
        "Charge or GL Amount",
        "DR or CR",
        "Cost Center",
        "Branch",
        " Charge Narration",
        "TaxGroup",
        "Tax Type",
        "SAC or HSN",
        "Taxcode1",
        "Taxcode1 Amt",
        "Taxcode2",
        "Taxcode2 Amt",
        "Taxcode3",
        "Taxcode3 Amt",
        "Taxcode4",
        "Taxcode4 Amt",
        "Avail Tax Credit",
        "LOB",
        "Ref Type",
        "Ref No",
        "Amount",
        "Start Date",
        "End Date",
        "WH Tax Code",
        "WH Tax Percentage",
        "WH Tax Taxable",
        "WH Tax Amount",
        "Round Off",
        "CC Code"
    ]
    if not all_details:
        log_callback("No data to write to CSV.")
        return False
    try:
        merged_details = []
        today_str = datetime.now().strftime("%d-%b-%Y")
        for row in all_details:
            merged_row = {**fixed_fields, **row}
            merged_row["Entry Date"] = today_str
            merged_row["Posting Date"] = today_str
            vendor_inv_date = merged_row.get("Vendor Inv Date", "")
            merged_row["Vendor Inv Date"] = convert_date_format(vendor_inv_date, "%d-%b-%Y") if vendor_inv_date else ""
            merged_row["Amount"] = row.get("Amount", "0")

            if row.get("is_transport"):
                # --- Transport Charges _ FCM(E) Block ---
                merged_row["Charge or GL Name"] = "Transport Charges _ FCM"
                merged_row["Tax Type"] = "Taxable"
                merged_row["TaxGroup"] = "GSTIN"
                merged_row["Avail Tax Credit"] = "100"
                merged_row["Taxcode1"] = "CGST"
                merged_row["Taxcode2"] = "SGST"
                merged_row["Taxcode1 Amt"] = str(round(float(row.get("Total Amount", "0")) * 0.06, 2))
                merged_row["Taxcode2 Amt"] = str(round(float(row.get("Total Amount", "0")) * 0.06, 2))
                merged_row["SAC or HSN"] = "996793"
                merged_row["Charge or GL Amount"] = row.get("Total Amount", "0")
                merged_row["Amount"] = row.get("Total Amount", "0")
                merged_row["WH Tax Taxable"] = row.get("Total Amount", "0")
                try:
                    merged_row["WH Tax Amount"] = str(round(float(merged_row["WH Tax Taxable"]) * 0.02, 2))
                except Exception:
                    merged_row["WH Tax Amount"] = "0"
            else:
                # --- CFS CHARGES (1) Block ---
                merged_row["Charge or GL Name"] = "CFS CHARGES (1)"
                merged_row["Tax Type"] = "Pure Agent"
                merged_row["TaxGroup"] = "GSTIN"
                merged_row["Avail Tax Credit"] = "No"
                merged_row["Taxcode1"] = ""
                merged_row["Taxcode2"] = ""
                merged_row["Taxcode1 Amt"] = ""
                merged_row["Taxcode2 Amt"] = ""
                merged_row["SAC or HSN"] = "996711"
                merged_row["Charge or GL Amount"] = row.get("Amount", "0")
                merged_row["WH Tax Taxable"] = row.get("WH Tax Taxable", "0")
                try:
                    merged_row["WH Tax Amount"] = str(round(float(merged_row["WH Tax Taxable"]) * 0.02, 2))
                except Exception:
                    merged_row["WH Tax Amount"] = "0"

            merged_row["Narration"] = f"BEING CHARGES PAID TO HASTI PETRO CHEMICAL A/C ADVICS {merged_row.get('Ref No', '')}"
            for col in column_order:
                if col not in merged_row:
                    merged_row[col] = ""
            merged_details.append(merged_row)
        with open(output_path, mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=column_order)
            writer.writeheader()
            for row in merged_details:
                writer.writerow({col: row.get(col, "") for col in column_order})
        log_callback(f"CSV successfully written to {output_path}")
        return True
    except Exception as e:
        log_callback(f"Failed to write CSV: {e}")
        return False

class DOInvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HASTI DO Invoice to CSV Converter")

        # Fullscreen
        try:
            self.root.state("zoomed")
        except:
            self.root.attributes("-fullscreen", True)

        self.root.configure(bg=BG_COLOR)

        # Variables
        self.pdf_paths = []
        self.job_register_path = None
        self.job_register = []
        self._logo_image = None

        # Setup Styles
        self._setup_styles()

        # Build UI
        self._create_widgets()

        # Logging Setup
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(text_handler)

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except:
            pass

        style.configure("Card.TLabelframe", background=CARD_BG, borderwidth=1, relief="solid")
        style.configure("Card.TLabelframe.Label", background=CARD_BG, foreground=TEXT_PRIMARY, font=("Segoe UI", 10, "bold"))

        style.configure("Modern.TButton", font=("Segoe UI", 9), padding=(14, 6), background=CARD_BG, borderwidth=1, relief="solid")
        style.map("Modern.TButton", background=[("active", "#F5F5F5"), ("pressed", "#EEEEEE")])

        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(20, 8), foreground="#FFFFFF", background=ACCENT, borderwidth=0)
        style.map("Accent.TButton", background=[("active", ACCENT_HOVER), ("pressed", ACCENT_HOVER), ("disabled", "#90CAF9")], foreground=[("disabled", "#FFFFFF")])

    def _create_widgets(self):
        # MAIN CONTAINER
        main_frame = tk.Frame(self.root, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ====== HEADER ======
        header_frame = tk.Frame(main_frame, bg=CARD_BG, pady=16, padx=24)
        header_frame.pack(fill=tk.X)
        tk.Frame(main_frame, bg=BORDER_COLOR, height=1).pack(fill=tk.X)

        # Logo (Left)
        logo_path = resource_path("logo.png")
        if HAS_PIL and os.path.isfile(logo_path):
            try:
                img = Image.open(logo_path)
                h = 40
                w = int(img.width * h / img.height)
                img = img.resize((w, h), Image.LANCZOS)
                self._logo_image = ImageTk.PhotoImage(img)
                tk.Label(header_frame, image=self._logo_image, bg=CARD_BG).pack(side=tk.LEFT)
            except Exception:
                tk.Label(header_frame, text="NAGARKOT", font=("Segoe UI", 12, "bold"), fg=ACCENT, bg=CARD_BG).pack(side=tk.LEFT)
        else:
            tk.Label(header_frame, text="NAGARKOT", font=("Segoe UI", 12, "bold"), fg=ACCENT, bg=CARD_BG).pack(side=tk.LEFT)

        # Centered Title
        tk.Label(
            header_frame, text="HASTI DO Invoice to CSV Converter",
            font=("Segoe UI", 16, "bold"), bg=CARD_BG, fg=TEXT_PRIMARY,
        ).place(relx=0.5, rely=0.3, anchor="center")

        tk.Label(
            header_frame, text="Parse DO Invoice PDFs and generate Logisys Upload CSV",
            font=("Segoe UI", 9), bg=CARD_BG, fg=TEXT_SECONDARY,
        ).place(relx=0.5, rely=0.75, anchor="center")

        # ====== BODY ======
        body = tk.Frame(main_frame, bg=BG_COLOR, padx=40, pady=30)
        body.pack(fill=tk.BOTH, expand=True)

        # --- File Selection Card ---
        file_card = ttk.LabelFrame(body, text="  Input Files  ", style="Card.TLabelframe", padding=20)
        file_card.pack(fill=tk.X, pady=(0, 20))

        file_inner = tk.Frame(file_card, bg=CARD_BG)
        file_inner.pack(fill=tk.BOTH, expand=True)

        self.pdf_status_label = tk.Label(file_inner, text="Invoice PDFs: Not Selected", fg=TEXT_SECONDARY, bg=CARD_BG, font=("Segoe UI", 9))
        self.pdf_status_label.pack(anchor=tk.W, pady=(0, 5))

        self.jobreg_status_label = tk.Label(file_inner, text="Job Register: Not Selected", fg=TEXT_SECONDARY, bg=CARD_BG, font=("Segoe UI", 9))
        self.jobreg_status_label.pack(anchor=tk.W, pady=(0, 15))

        btn_frame = tk.Frame(file_inner, bg=CARD_BG)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Select Invoice PDFs", command=self.select_pdf, style="Modern.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Select Job Register", command=self.select_job_register, style="Modern.TButton").pack(side=tk.LEFT)

        # --- Action Area ---
        action_frame = tk.Frame(body, bg=BG_COLOR)
        action_frame.pack(fill=tk.X, pady=(0, 20))

        self.process_button = ttk.Button(
            action_frame, text="\u25B6  Process & Generate CSV",
            command=self.process_files, style="Accent.TButton"
        )
        self.process_button.pack(side=tk.LEFT, padx=(0, 20))

        self.status_label = tk.Label(action_frame, text="Ready", fg=TEXT_SECONDARY, bg=BG_COLOR, font=("Segoe UI", 9))
        self.status_label.pack(side=tk.LEFT)

        # --- Log Card ---
        log_card = ttk.LabelFrame(body, text="  Processing Log  ", style="Card.TLabelframe", padding=15)
        log_card.pack(fill=tk.BOTH, expand=True)

        log_inner = tk.Frame(log_card, bg=CARD_BG)
        log_inner.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_inner, height=10, wrap=tk.WORD, state="disabled",
            bg=LOG_BG, fg=LOG_FG, font=("Consolas", 9),
            relief="flat", padx=10, pady=10,
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ====== FOOTER ======
        footer_frame = tk.Frame(main_frame, bg=CARD_BG, padx=24, pady=10)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Frame(main_frame, bg=BORDER_COLOR, height=1).pack(fill=tk.X, side=tk.BOTTOM)

        tk.Label(
            footer_frame, text="Nagarkot Forwarders Pvt. Ltd. \u00A9",
            fg=TEXT_SECONDARY, bg=CARD_BG, font=("Segoe UI", 8),
        ).pack(side=tk.LEFT)

        ttk.Button(footer_frame, text="Exit", command=self.root.destroy, style="Modern.TButton").pack(side=tk.RIGHT)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')}: {message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.root.update()

    def select_pdf(self):
        pdf_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not pdf_paths:
            self.log("No PDFs selected.")
            return
        self.pdf_paths = pdf_paths
        self.pdf_status_label.config(text=f"Invoice PDFs: {len(pdf_paths)} file(s) selected", fg=TEXT_PRIMARY)
        self.log(f"Selected {len(pdf_paths)} PDFs: {', '.join([os.path.basename(p) for p in pdf_paths])}")
        logger.info(f"Selected {len(pdf_paths)} PDFs")

    def select_job_register(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV or Excel files", "*.csv;*.xlsx;*.xls")])
        if not file_path:
            self.log("No job register selected.")
            return
        self.job_register_path = file_path
        self.jobreg_status_label.config(text=f"Job Register: {os.path.basename(file_path)}", fg=TEXT_PRIMARY)
        self.log(f"Selected job register: {os.path.basename(file_path)}")
        logger.info(f"Selected job register: {file_path}")
        self.load_job_register()

    def load_job_register(self):
        if not self.job_register_path:
            self.job_register = []
            return
        try:
            if self.job_register_path.endswith('.csv'):
                df = pd.read_csv(self.job_register_path, dtype=str)
            else:
                df = pd.read_excel(self.job_register_path, dtype=str)
            df.columns = [c.strip().lower() for c in df.columns]
            be_col = 'be no'
            job_col = 'job no'
            self.job_register = []
            for _, row in df.iterrows():
                self.job_register.append({
                    'be_no': str(row.get(be_col, '')).strip(),
                    'job_no': str(row.get(job_col, '')).strip(),
                })
            self.log(f"Loaded {len(self.job_register)} job register entries.")
        except Exception as e:
            self.log(f"Failed to load job register: {e}")
            self.job_register = []

    def match_job_no_by_be(self, be_no):
        be_no = str(be_no).strip()
        for entry in self.job_register:
            if entry.get('be_no', '') == be_no:
                return entry.get('job_no', 'No match found')
        return 'No match found'

    def process_files(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        logger.info("Starting file processing")

        if not self.pdf_paths:
            messagebox.showerror("Error", "No PDFs selected")
            self.log("No PDFs selected")
            return
        if not self.job_register_path or not hasattr(self, 'job_register') or not self.job_register:
            messagebox.showerror("Error", "No job register selected or loaded")
            self.log("No job register selected or loaded")
            return

        self.status_label.config(text="Processing...", fg=ACCENT)
        self.process_button.state(['disabled'])
        self.log("Starting processing")
        self.root.update()

        # Output directory
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(base_dir, "HASTI_Output")
        os.makedirs(output_dir, exist_ok=True)
        logger.info(f"Output directory: {output_dir}")

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        output_csv = os.path.join(output_dir, f"Hasti_{timestamp}.csv")
        logger.info(f"Output CSV: {output_csv}")

        if os.path.exists(output_csv):
            response = messagebox.askyesno(
                "File Exists",
                f"CSV file '{os.path.basename(output_csv)}' already exists. Overwrite?",
                parent=self.root
            )
            if not response:
                self.log("Cancelled overwrite.")
                self.status_label.config(text="Cancelled", fg=TEXT_SECONDARY)
                self.process_button.state(['!disabled'])
                return

        all_details = []
        for pdf_path in self.pdf_paths:
            self.log(f"Processing {os.path.basename(pdf_path)}")
            logger.info(f"Processing {pdf_path}")
            try:
                text, tables_data = extract_text_from_pdf(pdf_path, self.log)
                if not text:
                    self.log(f"Failed to extract text from {os.path.basename(pdf_path)}")
                    logger.error(f"Text extraction failed for {pdf_path}")
                    continue
                details_list = extract_invoice_details_with_regex(text, tables_data, self.log)
                if not details_list:
                    self.log(f"No valid data extracted from {os.path.basename(pdf_path)}")
                    logger.warning(f"No valid data extracted from {pdf_path}")
                    continue
                for details in details_list:
                    be_no = details.get("BOE No", "")
                    mapped_job_no = self.match_job_no_by_be(be_no)
                    details["Ref No"] = mapped_job_no
                    all_details.append(details)
                self.log(f"Processed {os.path.basename(pdf_path)}: {len(details_list)} records extracted")
                logger.info(f"Processed {pdf_path}: {len(details_list)} records extracted")
            except Exception as e:
                self.log(f"Failed to process {os.path.basename(pdf_path)}: {str(e)}")
                logger.error(f"Processing failed for {pdf_path}: {e}")
                continue

        if not all_details:
            self.status_label.config(text="No valid data extracted", fg=ERROR_RED)
            self.log("No valid data extracted from PDFs")
            messagebox.showerror("Error", "No valid data extracted from PDFs")
            self.process_button.state(['!disabled'])
            return

        if create_csv(all_details, output_csv, self.log):
            self.status_label.config(text="Completed Successfully", fg=SUCCESS_COLOR)
            self.log(f"CSV generated with {len(all_details)} records: {os.path.basename(output_csv)}")
            messagebox.showinfo("Success", f"CSV saved to {output_csv} with {len(all_details)} records")
        else:
            self.status_label.config(text="Failed", fg=ERROR_RED)
            self.log("Failed to generate CSV")
            messagebox.showerror("Error", "Failed to generate CSV")

        self.process_button.state(['!disabled'])

# Main
def main():
    try:
        logger.info("Starting HASTI DO Invoice Processor")
        root = tk.Tk()
        app = DOInvoiceApp(root)
        root.mainloop()
        logger.info("Application closed")
    except Exception as e:
        logger.error(f"Application error: {e}")
        messagebox.showerror("Error", f"Application error: {e}")

if __name__ == "__main__":
    main()
