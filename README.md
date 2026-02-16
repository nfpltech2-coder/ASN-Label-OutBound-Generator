# InBound & OutBound Data Generator

Automated tool for ASN data extraction and label generation for Nagarkot Forwarders Pvt Ltd.

## Tech Stack
- Python 3.10+
- Tkinter & CustomTkinter (GUI)
- pdfplumber (PDF Extraction)
- PyMuPDF (Label Generation)
- Pandas (Data Handling)

---

## Installation

### Clone / Extract
Copy the project files to your local directory.

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. **Create virtual environment**
   ```powershell
   python -m venv venv
   ```

2. **Activate (REQUIRED)**
   ```powershell
   venv\Scripts\activate
   ```

3. **Install dependencies**
   ```powershell
   pip install -r requirements.txt
   ```

4. **Run application**
   ```powershell
   python asn_gui_v2.py
   ```

---

### Build Executable (For Desktop Apps)

1. **Install PyInstaller** (Inside venv):
   ```powershell
   pip install pyinstaller
   ```

2. **Build using the included Spec file**:
   ```powershell
   pyinstaller asn_gui_v2.spec
   ```

3. **Locate Executable**:
   The application will be generated in the `dist/` folder.

---

## Notes
- **ALWAYS use virtual environment for Python.**
- The application launches in Full Screen mode (Press **Esc** to exit full screen).
- Assets like `logo.png` and `Label Quantity.xlsx` must be in the same directory as the script or bundled via spec file.
