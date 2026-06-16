"""
Quick local test: run with
  python test_pdf_parse.py /path/to/statement.pdf
"""
import sys, io, pdfplumber

PDF_PASSWORD = "SUDH3108"

path = sys.argv[1] if len(sys.argv) > 1 else "data/icici_31_statement.pdf"

with open(path, "rb") as f:
    raw = f.read()

print(f"PDF size: {len(raw):,} bytes")

with pdfplumber.open(io.BytesIO(raw), password=PDF_PASSWORD) as pdf:
    print(f"Pages: {len(pdf.pages)}")
    for i, page in enumerate(pdf.pages):
        print(f"\n--- Page {i+1} TEXT (first 800 chars) ---")
        txt = page.extract_text() or ""
        print(txt[:800])

        tables = page.extract_tables()
        print(f"\n--- Page {i+1}: {len(tables)} table(s) found ---")
        for ti, table in enumerate(tables):
            print(f"  Table {ti+1}: {len(table)} rows")
            for row in table[:5]:   # first 5 rows
                print("   ", row)
