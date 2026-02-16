import pdfplumber

with pdfplumber.open("Invoice.pdf") as pdf:
    text = pdf.pages[0].extract_text()
    print("--- FULL PAGE 1 TEXT ---")
    print(text)
    print("--- END PAGE 1 TEXT ---")
