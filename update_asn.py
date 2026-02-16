import pdfplumber
import re
import pandas as pd
import os
import sys
from datetime import datetime

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def load_csv_mapping(filename, code_col_preferred, name_col_preferred):
    # Use resource_path for loading CSVs bundled in the EXE
    filepath = resource_path(filename)
    if not os.path.exists(filepath):
        return []
    try:
        df = pd.read_csv(filepath, quotechar='"', skipinitialspace=True)
        # Robust column detection
        cols = [c.strip().strip('"') for c in df.columns]
        df.columns = cols
        
        target_code_col = code_col_preferred if code_col_preferred in cols else ("Consignee Code" if "Consignee Code" in cols else ("Supplier Code" if "Supplier Code" in cols else cols[1]))
        target_name_col = name_col_preferred if name_col_preferred in cols else ("Consignee Name" if "Consignee Name" in cols else ("Supplier Name" if "Supplier Name" in cols else cols[2]))
        
        mapping = []
        for _, row in df.iterrows():
            code = str(row[target_code_col]).strip().strip('"')
            name = str(row[target_name_col]).strip().strip('"')
            if name and code and name != "nan" and code != "nan":
                mapping.append((name.lower(), code))
        return mapping
    except Exception as e:
        print(f"Error loading {filename}: {e}")
        return []

def extract_all_data(pdf_path, doc_type="Invoice"):
    supplier_mapping = load_csv_mapping("Supplier Code.csv", "Supplier Code", "Supplier Name")
    consignee_mapping = load_csv_mapping("Consignee Code.csv", "Consignee Code", "Consignee Name")
    # Load postcode data as a list of dictionaries for robust searching
    postcode_data = []
    pc_filepath = resource_path("Ship to Address Code.csv")
    if os.path.exists(pc_filepath):
        try:
            pc_df = pd.read_csv(pc_filepath, quotechar='"', skipinitialspace=True)
            pc_df.columns = [c.strip().strip('"') for c in pc_df.columns]
            postcode_data = pc_df.to_dict('records')
        except Exception as e:
            print(f"Error loading postcode CSV: {e}")

    products = []
    po_number = "Unknown"
    po_date = "Unknown"
    found_code = "Unknown"
    found_name = ""
    
    with pdfplumber.open(pdf_path) as pdf:
        # 1. Extract Header Info
        header_text = ""
        for page in pdf.pages[:1]:
            header_text += page.extract_text() or ""
            header_text += " " + page.extract_text(layout=True) or ""
        
        header_text_lower = header_text.lower()

        # Extract PO Number
        po_label_match = re.search(r'Cust PO No\.\s*:\s*([A-Za-z0-9-]+)', header_text)
        if po_label_match:
            po_number = po_label_match.group(1)
        else:
            po_matches = re.findall(r'([0-9]{5,6}[A-Z][0-9]{6})', header_text)
            if po_matches:
                po_number = po_matches[0]
        
        # Extract PO Date
        date_patterns = [
            r'([A-Za-z]+ \d{1,2},?\s?\d{4})',       # December 20, 2025
            r'(\d{1,2}-[A-Za-z]+-\d{2,4})',       # 29-Jan-26 (Seen in 1835.pdf)
            r'(\d{4}-\d{2}-\d{2})',               # 2025-12-20
            r'(\d{2}/\d{2}/\d{4})'                # 20/12/2025
        ]
        
        for p in date_patterns:
            date_matches = re.findall(p, header_text)
            if date_matches:
                for raw_date in date_matches:
                    try:
                        clean_date = raw_date.replace(',', ', ').strip()
                        dt = pd.to_datetime(clean_date)
                        po_date = dt.strftime('%Y-%m-%d')
                        break
                    except:
                        continue
                if po_date != "Unknown":
                    break
        
        # Extract Invoice Number (Commonly requested to be used as DO No in OutBound)
        invoice_no = "Unknown"
        # Support common labels and characters including / and -
        inv_match = re.search(r'Invoice\s*(?:No\.|Number)\s*:\s*([A-Za-z0-9\/-]+)', header_text)
        if not inv_match:
            # Try without colon
            inv_match = re.search(r'Invoice\s*(?:No\.|Number)\s+([A-Za-z0-9\/-]+)', header_text)
            
        if inv_match:
            invoice_no = inv_match.group(1)
        else:
            # Special case for 1859-like format if header_text is messy
            inv_match_fallback = re.search(r'Invoice\s*Number\s*[:]\s*([^\s]+)', header_text)
            if inv_match_fallback:
                invoice_no = inv_match_fallback.group(1)

        # 2. Identify Supplier / Consignee
        if doc_type == "Invoice":
            # Search in Supplier Mapping (Priority: Longest Name)
            
            # Sub-Pass 0: Extract "Seller/From" specific section if available
            seller_match = re.search(r'(?:Seller|From|Shipper|Exporter)\s*(.*?)(?=\s*Consignee|\s*Billed To|\s*Description|\s*Invoice No|$)', header_text, re.DOTALL | re.IGNORECASE)
            seller_text = seller_match.group(1).lower() if seller_match else ""
            
            supplier_mapping.sort(key=lambda x: len(x[0]), reverse=True)
            
            # Priority 1: Check in specific Seller section
            if seller_text:
                for name, code in supplier_mapping:
                    if name and name in seller_text:
                        found_code = code
                        break
            
            # Priority 2: Check global header if not found
            if found_code == "Unknown":
                for name, code in supplier_mapping:
                    if name and name in header_text_lower:
                        found_code = code
                        break
            
            # Priority 3: Fallback for Japanese Suppliers via address (e.g. ADVICS JAPAN)
            if found_code == "Unknown":
                if "japan" in header_text_lower and ("kariya" in header_text_lower or "showa-cho" in header_text_lower):
                    found_code = "ADVICS-JAPAN"
        else:
            # Search in Consignee Mapping (Multi-pass logic)
            
            # Sub-Pass 0: Extract "Billed To" specific section
            billed_to_match = re.search(r'Billed To\s*(.*?)(?=\s*Shipped From|\s*Description|\s*Invoice No|$)', header_text, re.DOTALL | re.IGNORECASE)
            billed_to_text = billed_to_match.group(1).lower() if billed_to_match else ""
            
            # Pass 1: Match by Entity Name (Longest first for specificity)
            consignee_mapping.sort(key=lambda x: len(x[0]), reverse=True)
            
            # Priority 1: Check if any name exists in the "Billed To" block specifically
            if billed_to_text:
                for name, code in consignee_mapping:
                    if len(name) > 3 and name in billed_to_text:
                        found_name = name
                        found_code = code
                        # Specific User Override for Maruti Suzuki India Limited
                        if "maruti suzuki india limited" in name:
                            found_code = "C4000131"
                        break
            
            # Priority 2: Check global header if not found in Billed To
            if found_code == "Unknown":
                for name, code in consignee_mapping:
                    if len(name) > 3 and name in header_text_lower:
                        found_name = name
                        # Specific User Override for Maruti Suzuki India Limited
                        if "maruti suzuki india limited" in name:
                            found_code = "C4000131"
                        else:
                            found_code = code
                        break
            
            # Pass 2: Match by Participant Code
            if found_code == "Unknown":
                consignee_mapping.sort(key=lambda x: len(x[1]), reverse=True)
                for name, code in consignee_mapping:
                    if len(code) > 2 and code.lower() in header_text_lower:
                        found_code = code
                        found_name = name
                        break
            
            # Pass 3: Match by Address Lines (Fallback for "Unknown" consignees)
            if found_code == "Unknown" and postcode_data:
                for entry in postcode_data:
                    # Check Line Address 1 and 2
                    line1 = str(entry.get('Line Address 1', '')).strip().lower()
                    line2 = str(entry.get('Line Address 2', '')).strip().lower()
                    
                    # We need a reasonably long match to avoid false positives
                    if (len(line1) > 10 and line1 in header_text_lower) or \
                       (len(line2) > 10 and line2 in header_text_lower):
                        found_code = str(entry.get('Consignee Code', 'Unknown')).strip()
                        # Also attempt to find the name in the Consignee Mapping if possible,
                        # otherwise use the code as name for reference
                        found_name = ""
                        for mapping_name, mapping_code in consignee_mapping:
                            if mapping_code == found_code:
                                found_name = mapping_name
                                break
                        break
 
        if found_code == "Unknown":
            # Fallback for "ADVICS" via address
            if "japan" in header_text_lower and "kariya" in header_text_lower:
                found_code = "ADVICS-JAPAN"
 
        # 3. Extract Products
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            
            lines = text.split('\n')
            for line in lines:
                # Optimized regex: Capture item code, optionally skip 8-digit HSN, capture qty (decimal/commas)
                match = re.search(r'(\d{6}-\d{5})\s+(?:\d{8}\s+)?([\d,.]+)', line)
                if match:
                    p_code = match.group(1)
                    qty_str = match.group(2).replace(',', '')
                    try:
                        qty_val = float(qty_str)
                        # PCS are typically whole numbers; keep as int if possible
                        qty = int(qty_val) if qty_val.is_integer() else qty_val
                    except:
                        continue
                    
                    if doc_type == "Invoice":
                        products.append({
                            'storerCode': 'AIPL',
                            'warehouseCode': 'NFKD',
                            'poNumber': po_number,
                            'poDate': po_date,
                            'supplierCode': found_code,
                            'otherReference': '',
                            'productCode': p_code,
                            'quantity': int(qty),
                            'uomCode': 'PCS',
                            'fileName': os.path.basename(pdf_path)
                        })
                    else: # OutBound
                        # Robust Postcode Lookup
                        postcode = ""
                        # Step 1: Lookup by Consignee Code
                        for entry in postcode_data:
                            if str(entry.get('Consignee Code', '')).strip().lower() == found_code.lower():
                                postcode = str(entry.get('Post Code', '')).strip()
                                break
                        
                        # Step 2: Fallback to Lookup by Matched Name in Address Lines
                        if not postcode and found_name:
                            for entry in postcode_data:
                                addr1 = str(entry.get('Line Address 1', '')).lower()
                                addr2 = str(entry.get('Line Address 2', '')).lower()
                                if found_name.lower() in addr1 or found_name.lower() in addr2:
                                    postcode = str(entry.get('Post Code', '')).strip()
                                    break
                        
                        # Use Invoice No as DO Number for OutBound if found, else fallback to po_number
                        final_do_no = invoice_no if invoice_no != "Unknown" else po_number
                        
                        products.append({
                            'storerCode': 'AIPL',
                            'warehouseCode': 'NFKD',
                            'doNumber': final_do_no,
                            'consigneeCode': found_code,
                            'shipToAddressPostCode': postcode,
                            'requiredDate': po_date,
                            'otherReference': '',
                            'productCode': p_code,
                            'quantity': int(qty),
                            'uomCode': 'PCS',
                            'fileName': os.path.basename(pdf_path)
                        })
        
    return products

def update_excel(excel_path, new_data, doc_type="Invoice"):
    if not new_data:
        return False
    try:
        if doc_type == "Invoice":
            cols = ['storerCode', 'warehouseCode', 'poNumber', 'poDate', 'supplierCode', 'otherReference', 'productCode', 'quantity', 'uomCode']
        else:
            cols = ['storerCode', 'warehouseCode', 'doNumber', 'consigneeCode', 'shipToAddressPostCode', 'requiredDate', 'otherReference', 'productCode', 'quantity', 'uomCode']
            
        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path)
            # Ensure we have all columns
            for c in cols:
                if c not in df.columns:
                    df[c] = ""
        else:
            df = pd.DataFrame(columns=cols)

        new_df = pd.DataFrame(new_data)
        # Filter new_df to only include columns that exist in the template
        valid_cols = [c for c in cols if c in new_df.columns]
        new_df = new_df[valid_cols]
        
        updated_df = pd.concat([df, new_df], ignore_index=True)
        # Final column order insurance
        updated_df = updated_df[cols]
        
        updated_df.to_excel(excel_path, index=False)
        return True
    except Exception as e:
        print(f"Error updating Excel: {e}")
        return False

if __name__ == "__main__":
    # Test script mode
    data = extract_all_data("Invoice.pdf", "Invoice")
    if data:
        update_excel("ASN Data.xlsx", data, "Invoice")
        print("Invoice Done.")
    
    data_pl = extract_all_data("Packing List.pdf", "Packing List")
    if data_pl:
        update_excel("NAGARKOT_ORDER_UPLOAD_TEMPLATE.xlsx", data_pl, "Packing List")
        print("Packing List Done.")
