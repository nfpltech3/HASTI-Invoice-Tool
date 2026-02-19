# HASTI Invoice to CSV Tool

A Python desktop application for converting HASTI Petro Chemical DO invoice PDFs into Logisys-compatible CSV files.

## Tool Included

| App | Purpose |
| --- | --- |
| `HASTI_Invoice_to_CSV.py` | Parses DO invoice PDFs and generates purchase CSV |

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

4. Run application
```bash
python HASTI_Invoice_to_CSV.py
```

---

## Build Executable

1. Install PyInstaller (Inside venv):
```bash
pip install pyinstaller
```

2. Build using included Spec file:
```bash
pyinstaller HASTI_Invoice_to_CSV.spec
```

3. Locate Executable:
The executable will be generated in the `dist/` folder.

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit venv, node_modules, dist, or build folders.
- Output CSVs are saved in `HASTI_Output/`.
- Run and test before pushing.
