# HASTI Invoice & Ledger to CSV Tools

Two Python desktop apps for converting HASTI Petro Chemical invoices and ledger reports into Logisys-compatible CSV files.

## Tools Included

| App | Purpose |
| --- | --- |
| `HASTI_Invoice_to_CSV.py` | Parses DO invoice PDFs and generates purchase CSV |
| `Ledger_to_CSV.py` | Converts ledger Excel reports with Job Register lookup |

## Tech Stack
- Python 3.11+
- Tkinter (GUI)
- Pandas / OpenPyXL (Data)
- pdfplumber (PDF Parsing)
- Pillow (Logo rendering)

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create virtual environment
```bash
python -m venv venv
```

2. Activate (REQUIRED)

Windows:
```cmd
venv\Scripts\activate
```

Mac/Linux:
```bash
source venv/bin/activate
```

3. Install dependencies
```bash
pip install -r requirements.txt
```

4. Run applications
```bash
python HASTI_Invoice_to_CSV.py
python Ledger_to_CSV.py
```

---

## Build Executables

1. Install PyInstaller (Inside venv):
```bash
pip install pyinstaller
```

2. Build using included Spec files:
```bash
pyinstaller HASTI_Invoice_to_CSV.spec
pyinstaller Ledger_to_CSV.spec
```

3. Locate Executables:
Both will be generated in the `dist/` folder.

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit venv, node_modules, dist, or build folders.
- Output CSVs are saved in `HASTI_Output/` (Invoice) or `Kale Output/` (Ledger).
- Run and test before pushing.
