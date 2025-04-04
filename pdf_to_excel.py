
import fitz  # PyMuPDF
import pandas as pd
import re
from collections import defaultdict

# PDF input file path
pdf_file = r"C:\Users\hp\pdf_reader\sample_tables.pdf"

# Output Excel path
output_excel = r"C:\Users\hp\pdf_reader\sample_tables.xlsx"

def clean_text(text):
    """Removes illegal characters for Excel from a string."""
    if not isinstance(text, str):
        return text
    return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', text)

def extract_tables_from_pdf(pdf_path):
    """Extracts only tabular data based on column layout from the PDF."""
    doc = fitz.open(pdf_path)
    all_tables = []

    for page_num, page in enumerate(doc, start=1):
        text_dict = page.get_text("dict")
        words = []
        for block in text_dict["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    words.append({
                        "x": span["bbox"][0],
                        "y": span["bbox"][1],
                        "text": span["text"].strip()
                    })

        # Group into rows
        rows = defaultdict(list)
        for word in words:
            y_key = round(word["y"] / 2.5)  # tweak if needed
            rows[y_key].append((word["x"], word["text"]))

        table = []
        for y_key in sorted(rows.keys()):
            line = sorted(rows[y_key], key=lambda x: x[0])
            row_texts = [cell[1] for cell in line if cell[1]]

            # Consider it a "table row" only if it has multiple columns
            if len(row_texts) >= 3:
                table.append(row_texts)

        if table:
            all_tables.append((f"Page_{page_num}", table))

    return all_tables

def save_tables_to_excel(tables, output_path):
    """Saves extracted tables into an Excel file with one sheet per page."""
    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    for sheet_name, table in tables:
        cleaned_table = [[clean_text(cell) for cell in row] for row in table]
        df = pd.DataFrame(cleaned_table)
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False, header=False)
    writer.close()

# -------- Run the tool --------
if __name__ == "__main__":
    print("Extracting only tables from:", pdf_file)
    tables = extract_tables_from_pdf(pdf_file)
    if tables:
        save_tables_to_excel(tables, output_excel)
        print(f"\n✅ Tables successfully extracted and saved to: {output_excel}")
    else:
        print("\n⚠️ No tables found in the PDF.")
