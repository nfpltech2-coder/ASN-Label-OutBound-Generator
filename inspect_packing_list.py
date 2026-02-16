import pdfplumber

with pdfplumber.open("Packing List.pdf") as pdf:
    # First page full text
    text = pdf.pages[0].extract_text()
    print("--- Page 1 Text ---")
    print(text)
    
    # Try with layout
    text_layout = pdf.pages[0].extract_text(layout=True)
    print("\n--- Page 1 Layout ---")
    print(text_layout[:2000])

    # Tables
    tables = pdf.pages[0].extract_tables()
    print("\n--- Page 1 Tables ---")
    for table in tables:
        for row in table:
            print(row)
