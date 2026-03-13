"""Debug script to inspect raw text from Receipt Slip PDFs."""
import pdfplumber
import sys

pdfs = [
    r"c:\Projects\Backup\RECEIPT SLIP BIPL INVOICE NO - AN0250014165.pdf",
    r"c:\Projects\Backup\Receipt Slip ATI INVOICE NO - 910111300.pdf",
    r"c:\Projects\Backup\Receipt Slip ADVICS CO. LTD JAPAN INVOICE NO -12601Q100091.pdf",
]

output_lines = []
for target in pdfs:
    output_lines.append(f"\n{'='*60}")
    output_lines.append(f"FILE: {target}")
    output_lines.append('='*60)
    with pdfplumber.open(target) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            output_lines.append(f"\n--- PAGE {i+1} TEXT ---")
            output_lines.append(text)

out_path = r"c:\Projects\Backup\receipt_slip_debug.txt"
with open(out_path, "w", encoding="utf-8") as f:
    f.write("\n".join(output_lines))

print(f"Output written to {out_path}")
