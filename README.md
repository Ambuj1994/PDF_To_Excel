# PDF_To_Excel
Python code for extracting tables from PDF and storing in excel file

PDF Table Extractor Tool
=========================

Description:
------------
This Python tool extracts tabular data from scanned or digital PDF files using PyMuPDF (fitz)
and exports the data to an Excel (.xlsx) file. It detects tables based on layout positions,
making it useful even when traditional PDF table extractors fail.

Main Script:
------------
pdf_to_excel.py

Features:
---------
- Detects tables from scanned or native PDFs using text positioning.
- Handles multi-page PDFs and writes one sheet per page.
- Automatically cleans text for Excel compatibility.
- Simple and customizable code for tweaking detection sensitivity.

Usage Instructions:
-------------------
1. Open the `pdf_to_excel.py` file.

2. Set the following file paths in the script:
   - `pdf_file`: Path to your input PDF file (e.g., "C:\Users\hp\pdf_reader\foo.pdf") //Subject to change
   - `output_excel`: Desired output Excel path (e.g., "C:\Users\hp\pdf_reader\foo.xlsx")

3. Run the script:
   ```bash
   python pdf_to_excel.py
   ```

4. If tables are found, they will be saved to an Excel file with one sheet per PDF page.

Dependencies:
-------------
Ensure the following Python packages are installed:

- Python 3.7 or higher
- PyMuPDF (fitz)
- pandas
- openpyxl

To install dependencies, run:
```bash
pip install pymupdf pandas openpyxl
```


Submission Checklist:
---------------------
✔ Source code file: pdf_to_excel.py
✔ README.txt file with documentation
✔ Example input PDF (e.g., foo.pdf)
✔ Output Excel file (e.g., foo.xlsx)
✔ Demo video or slides (optional for presentation)

Contact:
--------
For any issues or feedback, please contact the developer or evaluator.


