import fitz  # PyMuPDF
import os
import pandas as pd
from datetime import datetime
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def load_quantity_mapping():
    mapping = {}
    try:
        file_path = resource_path("Label Quantity.xlsx")
        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
            # Expected columns: "Part Code", "PLT Quantity"
            # Normalize column names just in case
            df.columns = [c.strip() for c in df.columns]
            for _, row in df.iterrows():
                code = str(row.get("Part Code", "")).strip()
                qty = row.get("PLT Quantity", 0)
                if code:
                    mapping[code] = str(qty)
    except Exception as e:
        print(f"Error loading Label Quantity.xlsx: {e}")
    return mapping

def generate_label_pdf(data_list, output_path, grn_date=None):
    """
    Generates a PDF with labels for each item in data_list.
    Label size: 10cm x 7.5cm (approx 283 x 213 points).
    """
    # 10cm x 7.5cm (Landscape)
    width, height = 283, 213
    doc = fitz.open()

    if not grn_date:
        grn_date = datetime.now().strftime("%d-%b-%Y")

    for item in data_list:
        page = doc.new_page(width=width, height=height)
        
        # Data extraction
        po_number = str(item.get('poNumber', 'Unknown'))
        product_code = str(item.get('productCode', 'Unknown'))
        quantity = str(item.get('quantity', '0'))
        total_quantity = str(item.get('total_quantity', '0'))
        description = str(item.get('description', ''))
        
        # --- Draw Text ---
        def draw_centered_text(text, y, font_size, font_name="Helvetica", bold=False):
            font = f"{font_name}-Bold" if bold else font_name
            try:
                text_len = fitz.get_text_length(text, fontname=font, fontsize=font_size)
                x = (width - text_len) / 2
                page.insert_text((x, y + font_size), text, fontsize=font_size, fontname=font)
            except Exception as e:
                page.insert_text((10, y + font_size), text, fontsize=font_size, fontname=font)

        # 1. Invoice Number (Top)
        page.insert_text((10, 15), "Invoice No.", fontsize=10, fontname="Helvetica")
        draw_centered_text(po_number, 20, 36, bold=True)
        
        # Separator Line
        page.draw_line((0, 65), (width, 65), color=(0, 0, 0), width=1.5)

        # 2. Product Code (Middle)
        page.insert_text((10, 80), "Part Code", fontsize=10, fontname="Helvetica")
        draw_centered_text(product_code, 85, 40, bold=True)
        
        # Separator Line
        page.draw_line((0, 135), (width, 135), color=(0, 0, 0), width=1.2)

        # 3. Quantity (Bottom Left/Middle)
        page.insert_text((10, 150), "QTY:", fontsize=10, fontname="Helvetica-Bold")
        qty_display = f"{quantity} / {total_quantity}"
        page.insert_text((10, 175), qty_display, fontsize=24, fontname="Helvetica-Bold")
        
        # 4. GRN Date (Bottom Right)
        # We'll make this more centrally aligned with description below
        page.insert_text((width - 90, 150), "GRN Date:", fontsize=10, fontname="Helvetica")
        page.insert_text((width - 90, 170), grn_date, fontsize=12, fontname="Helvetica-Bold")

        # 5. Part Description (Very Bottom)
        # Using a small font as requested
        desc_text = description[:50] # Limit length just in case
        draw_centered_text(desc_text, 185, 10, font_name="Helvetica", bold=False)

    try:
        doc.save(output_path)
        doc.close()
        return True
    except Exception as e:
        print(f"Error saving PDF: {e}")
        return False
