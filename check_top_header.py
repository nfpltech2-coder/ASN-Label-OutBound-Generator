import pdfplumber

with pdfplumber.open("Invoice.pdf") as pdf:
    page = pdf.pages[0]
    # Very top of the page
    top_region = (0, 0, page.width, 100)
    text = page.within_bbox(top_region).extract_text()
    print("--- TOP 100 PIXELS TEXT ---")
    print(text)
    
    # Try with extract_words and full layout
    words = page.within_bbox(top_region).extract_words()
    print("--- TOP 100 PIXELS WORDS ---")
    for w in words:
        print(w['text'], end=" ")
    print("\n")
