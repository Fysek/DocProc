# DocProc

Employee Document Manager CLI
A powerful, modular Command Line Interface (CLI) application built in Python to automate and manage employee documentation. It validates file completeness, standardizes naming conventions, generates ZIP packages, and exports status reports to Excel.

🚀 Features
Validation & Standardization (check): Verifies if each employee has the "Big Five" required documents (BHP, Medical Certificate, ID, Photo, G2E). Automatically standardizes filenames to the First_Last_Document.ext format.

Automated Packaging (pack): Reads a target list of employees from a JSON file and generates a .zip archive containing their documents (only if the documentation is complete).

Reporting (excel): Generates an .xlsx spreadsheet summarizing the documentation status (present/missing) for all employees. Note: Data extraction (OCR) for expiration dates is planned for future releases.

PDF Compression (compress): Scans the database for PDF files exceeding 1MB. Note: Actual compression logic is planned for future releases.

📁 Expected Directory Structure
The application expects a main database folder containing subfolders named after each employee. Inside these subfolders are the actual document files (PDF, JPG, PNG).

Plaintext
Database_Folder/
├── Jan Kowalski/
│   ├── bhp.pdf
│   ├── lekarskie.jpg
│   ├── dowod.png
│   ├── foto.jpg
│   └── g2e.pdf
├── Adam Nowak/
│   └── ...
🛠️ Installation & Prerequisites
Ensure you have Python 3.7+ installed.

Install the required dependencies (currently only openpyxl for Excel generation):

Bash
pip install openpyxl
💻 Usage
The application uses subcommands. You must always provide the path to your main employee database folder using the --path argument.

1. Check & Standardize Documents
Validates completeness and renames files to the standard format.

Bash
python main.py check --path "C:\Path\To\Database"
2. Generate ZIP Packages
Creates ZIP archives for specific employees based on a JSON list.

Bash
python main.py pack --path "C:\Path\To\Database" workers.json
workers.json Template:

JSON
[
  "Jan Kowalski",
  "Adam Nowak"
]
3. Generate Excel Report
Creates raport_dokumentow.xlsx in the current directory with document statuses.

Bash
python main.py excel --path "C:\Path\To\Database"
4. Find PDFs for Compression
Lists all PDF files larger than 1MB.

Bash
python main.py compress --path "C:\Path\To\Database"
🔮 Future Roadmap
OCR Integration: Implement optical character recognition (e.g., pytesseract, pdfplumber) in the excel module to automatically extract expiration dates from medical and BHP certificates.

Active PDF Compression: Implement Ghostscript or pikepdf in the compress module to actively shrink large files.