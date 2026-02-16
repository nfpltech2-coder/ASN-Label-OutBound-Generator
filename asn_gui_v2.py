import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
import pandas as pd
import os
from datetime import datetime
from update_asn import extract_all_data, update_excel
from label_utils import generate_label_pdf, load_quantity_mapping
import sys
import io
import re
from PIL import Image, ImageTk
try:
    import win32com.client
except ImportError:
    win32com = None

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def load_product_quantity_mapping():
    """Load product code to pallet quantity mapping from Label Quantity.xlsx"""
    try:
        file_path = resource_path("Label Quantity.xlsx")
        if not os.path.exists(file_path):
            file_path = "Label Quantity.xlsx"
        
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        
        mapping = {}
        for _, row in df.iterrows():
            code = str(row.get("Part Code", row.iloc[0])).strip()
            qty = row.get("PLT Quantity", row.iloc[1])
            desc = row.get("Part Description", "")
            if code and code != 'nan':
                mapping[code] = {'qty': qty, 'desc': desc}
        return mapping
    except Exception as e:
        print(f"Error loading Label Quantity.xlsx: {e}")
        return {}

class ASNGeneratorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Nagarkot Brand Palette ---
        self.BRAND_COLORS = {
            "primary": "#1F3F6E",
            "accent": "#D8232A",
            "bg": "#F4F6F8",
            "white": "#FFFFFF",
            "text": "#1E1E1E",
            "muted": "#6B7280",
            "border": "#E5E7EB",
            "hover": "#2A528F"
        }

        self.title("Nagarkot Forwarders Pvt Ltd - ASN_Label_OutBound Generator")
        self.state('zoomed') # Start maximized but with decorations (unless we want to hide them)
        # If user wants a custom look without standard title bar but with buttons, 
        # normally we'd use overrideredirect(True), but let's stick to standard behavior 
        # for high safety unless they want a custom title bar.
        # Actually, let's add custom control buttons to the header for a premium feel.
        
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.configure(fg_color=self.BRAND_COLORS["bg"])

        self.extracted_data = [] # Stores ALL loaded products
        self.pdf_paths = []      # List of currently loaded PDF paths
        self.current_preview_file = None
        self.email_pdf_paths = []
        self.doc_type = ctk.StringVar(value="Invoice")
        self.entry_mode = ctk.StringVar(value="PDF")  # "PDF", "Manual", or "Email"
        
        # Load product quantity mapping
        self.product_qty_mapping = load_product_quantity_mapping()
        self.product_codes = sorted(list(self.product_qty_mapping.keys()))
        
        # Excel files
        self.excel_invoice = "InBound ASN Register.xlsx"
        self.excel_packing = "OutBound Order Register.xlsx"

        # --- Layout ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) # Body area

        self.create_header()
        
        # Main container for body to allow flexible layout
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=1, column=0, sticky="nsew", padx=40, pady=20)
        self.main_container.grid_columnconfigure(0, weight=1)

        # Mode Selection (InBound/OutBound)
        self.mode_frame = ctk.CTkFrame(self.main_container, fg_color=self.BRAND_COLORS["white"])
        self.mode_frame.grid(row=0, column=0, padx=0, pady=(0,10), sticky="ew")
        
        self.radio_invoice = ctk.CTkRadioButton(self.mode_frame, text="InBound", variable=self.doc_type, value="Invoice", command=self.on_mode_change)
        self.radio_invoice.pack(side="left", padx=40, pady=15)
        
        self.radio_packing = ctk.CTkRadioButton(self.mode_frame, text="OutBound", variable=self.doc_type, value="Packing List", command=self.on_mode_change, text_color=self.BRAND_COLORS["text"], fg_color=self.BRAND_COLORS["primary"])
        self.radio_packing.pack(side="left", padx=40, pady=15)

        # Entry Mode Selection (PDF Upload / Manual Entry) - Only for InBound
        self.entry_mode_frame = ctk.CTkFrame(self.main_container, fg_color=self.BRAND_COLORS["white"])
        
        self.radio_pdf = ctk.CTkRadioButton(self.entry_mode_frame, text="PDF Upload", variable=self.entry_mode, value="PDF", command=self.on_entry_mode_change, text_color=self.BRAND_COLORS["text"], fg_color=self.BRAND_COLORS["primary"])
        self.radio_pdf.pack(side="left", padx=40, pady=10)
        
        self.radio_manual = ctk.CTkRadioButton(self.entry_mode_frame, text="Manual Entry (Labels)", variable=self.entry_mode, value="Manual", command=self.on_entry_mode_change, text_color=self.BRAND_COLORS["text"], fg_color=self.BRAND_COLORS["primary"])
        self.radio_manual.pack(side="left", padx=40, pady=10)
        
        self.radio_email = ctk.CTkRadioButton(self.entry_mode_frame, text="Email Draft", variable=self.entry_mode, value="Email", command=self.on_entry_mode_change, text_color=self.BRAND_COLORS["text"], fg_color=self.BRAND_COLORS["primary"])
        self.radio_email.pack(side="left", padx=40, pady=10)

        # PDF Upload Section
        self.upload_frame = ctk.CTkFrame(self.main_container, fg_color=self.BRAND_COLORS["white"])
        
        self.btn_upload = ctk.CTkButton(self.upload_frame, text="Upload PDF(s)", command=self.upload_pdf, fg_color=self.BRAND_COLORS["primary"], hover_color=self.BRAND_COLORS["hover"])
        self.btn_upload.pack(side="left", padx=20, pady=20)

        self.btn_clear_pdfs = ctk.CTkButton(self.upload_frame, text="Clear All Files", command=self.clear_all_pdfs, fg_color="orange")
        self.btn_clear_pdfs.pack(side="left", padx=10, pady=20)

        self.lbl_status = ctk.CTkLabel(self.upload_frame, text="No files selected", font=ctk.CTkFont(slant="italic"), text_color=self.BRAND_COLORS["muted"])
        self.lbl_status.pack(side="left", padx=20)

        # Manual Entry Section
        self.manual_frame = ctk.CTkFrame(self.main_container, fg_color=self.BRAND_COLORS["white"])
        
        # GRN Date Field
        ctk.CTkLabel(self.manual_frame, text="GRN Date:", text_color=self.BRAND_COLORS["text"]).pack(side="left", padx=(20, 5), pady=10)
        self.grn_date_manual = ctk.CTkEntry(self.manual_frame, width=120)
        self.grn_date_manual.pack(side="left", padx=5, pady=10)
        self.grn_date_manual.insert(0, datetime.now().strftime("%d-%b-%Y"))

        # Buttons for manual entry
        self.btn_add_row = ctk.CTkButton(self.manual_frame, text="Add Row", command=self.add_manual_row, fg_color="green")
        self.btn_add_row.pack(side="left", padx=10, pady=10)
        
        self.btn_delete_row = ctk.CTkButton(self.manual_frame, text="Delete Selected", command=self.delete_selected_row, fg_color=self.BRAND_COLORS["accent"])
        self.btn_delete_row.pack(side="left", padx=10, pady=10)
        
        self.btn_clear_all = ctk.CTkButton(self.manual_frame, text="Clear All", command=self.clear_all_rows, fg_color="orange")
        self.btn_clear_all.pack(side="left", padx=10, pady=10)
        
        # Email Section
        self.email_frame = ctk.CTkFrame(self.main_container, fg_color=self.BRAND_COLORS["white"])
        
        self.btn_upload_email = ctk.CTkButton(self.email_frame, text="Select PDF Files", command=self.upload_email_pdfs, fg_color=self.BRAND_COLORS["primary"], hover_color=self.BRAND_COLORS["hover"])
        self.btn_upload_email.pack(side="left", padx=20, pady=20)
        
        self.lbl_email_status = ctk.CTkLabel(self.email_frame, text="No files selected", font=ctk.CTkFont(slant="italic"), text_color=self.BRAND_COLORS["muted"])
        self.lbl_email_status.pack(side="left", padx=20)
        
        self.btn_clear_email_files = ctk.CTkButton(self.email_frame, text="Clear Files", command=self.clear_email_files, fg_color="orange", width=100)
        self.btn_clear_email_files.pack(side="left", padx=10, pady=20)

        # Supplier and Invoice Fields for Email
        self.email_inputs_frame = ctk.CTkFrame(self.main_container, fg_color=self.BRAND_COLORS["white"])
        
        ctk.CTkLabel(self.email_inputs_frame, text="Supplier Name:", text_color=self.BRAND_COLORS["text"]).pack(side="left", padx=(20, 5), pady=10)
        self.email_supplier_entry = ctk.CTkEntry(self.email_inputs_frame, width=200)
        self.email_supplier_entry.pack(side="left", padx=5, pady=10)
        
        ctk.CTkLabel(self.email_inputs_frame, text="Invoice No:", text_color=self.BRAND_COLORS["text"]).pack(side="left", padx=(20, 5), pady=10)
        self.email_invoice_entry = ctk.CTkEntry(self.email_inputs_frame, width=150)
        self.email_invoice_entry.pack(side="left", padx=5, pady=10)

        self.btn_create_draft = ctk.CTkButton(self.email_frame, text="Create Outlook Draft", command=self.create_outlook_draft, fg_color="green")
        self.btn_create_draft.pack(side="left", padx=20, pady=20)

        # Preview/Table Section
        self.preview_frame = ctk.CTkFrame(self.main_container, fg_color=self.BRAND_COLORS["white"])
        self.preview_frame.grid_columnconfigure(0, weight=0) # Sidebar
        self.preview_frame.grid_columnconfigure(1, weight=1) # Main Preview
        self.preview_frame.grid_rowconfigure(0, weight=1)

        # File List Sidebar (only for PDF mode)
        self.file_sidebar = ctk.CTkScrollableFrame(self.preview_frame, width=200, label_text="Loaded Invoices", label_text_color=self.BRAND_COLORS["primary"])
        self.file_btns = {} # filename -> button widget

        # PDF Preview (Textbox)
        self.txt_preview = ctk.CTkTextbox(self.preview_frame, wrap="none", font=("Consolas", 12), fg_color=self.BRAND_COLORS["white"], text_color=self.BRAND_COLORS["text"])
        
        # Manual Entry Table (Widget-based grid in ScrollableFrame)
        self.create_manual_table()

        # Export Section
        self.export_frame = ctk.CTkFrame(self.main_container, fg_color=self.BRAND_COLORS["white"])
        self.export_frame.grid(row=5, column=0, padx=0, pady=20, sticky="ew")

        self.btn_save_excel = ctk.CTkButton(self.export_frame, text="Save to Excel", state="disabled", command=self.save_excel, fg_color=self.BRAND_COLORS["primary"], hover_color=self.BRAND_COLORS["hover"])
        self.btn_save_excel.pack(side="left", padx=20, pady=20)
        
        self.btn_download_txt = ctk.CTkButton(self.export_frame, text="Download (Tab Delimited)", state="disabled", command=self.download_txt, fg_color=self.BRAND_COLORS["primary"], hover_color=self.BRAND_COLORS["hover"])
        self.btn_download_txt.pack(side="left", padx=20, pady=20)

        self.btn_download_labels = ctk.CTkButton(self.export_frame, text="Download Labels", state="disabled", command=self.download_labels, fg_color="green")
        self.btn_download_labels.pack(side="left", padx=20, pady=20)

        self.create_footer()

        # Initial UI state
        self.on_entry_mode_change()
        self.on_mode_change()

    def create_header(self):
        """Create the Nagarkot branded header with logo"""
        self.header_frame = ctk.CTkFrame(self, fg_color=self.BRAND_COLORS["white"], height=60, corner_radius=0)
        self.header_frame.grid(row=0, column=0, sticky="ew")
        self.header_frame.grid_propagate(False)

        # Logo - Positioned left
        try:
            logo_path = resource_path("logo.png")
            img = Image.open(logo_path)
            # Aspect ratio resize to height 25 for better visibility
            w, h = img.size
            new_h = 25
            new_w = int((new_h / h) * w)
            img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img)
            self.logo_label = tk.Label(self.header_frame, image=self.logo_img, bg=self.BRAND_COLORS["white"])
            self.logo_label.pack(side="left", padx=(30, 10), pady=15)
        except Exception as e:
            print(f"Logo load error: {e}")
            self.logo_label = ctk.CTkLabel(self.header_frame, text="NAGARKOT", font=("Arial", 14, "bold"), text_color=self.BRAND_COLORS["primary"])
            self.logo_label.pack(side="left", padx=30, pady=15)

    def create_footer(self):
        """Create the Nagarkot branded footer"""
        self.footer_frame = ctk.CTkFrame(self, fg_color="transparent", height=30)
        self.footer_frame.grid(row=3, column=0, sticky="ew")
        
        self.footer_label = ctk.CTkLabel(self.footer_frame, text="Nagarkot Forwarders Pvt Ltd. ©", font=("Arial", 10), text_color=self.BRAND_COLORS["muted"])
        self.footer_label.pack(side="left", padx=40, pady=5)

    def create_manual_table(self):
        """Create the manual entry area with ScrollableFrame and headers"""
        # Container for table - Fixed to light color for readability
        self.manual_scroll_frame = ctk.CTkScrollableFrame(self.preview_frame, fg_color=self.BRAND_COLORS["white"])
        
        # Headers
        inv_label = "Invoice No." if self.doc_type.get() == "Invoice" else "DO Number"
        headers = ["Product Code", inv_label, "Total Quantity", "Pallet Quantity", ""]
        header_frame = ctk.CTkFrame(self.manual_scroll_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, columnspan=5, sticky="ew", padx=5, pady=5)
        
        # Set column weights
        for i in range(4):
            self.manual_scroll_frame.grid_columnconfigure(i, weight=1)
        self.manual_scroll_frame.grid_columnconfigure(4, weight=0) # Delete button column

        self.manual_header_widgets = []
        for col, text in enumerate(headers):
            lbl = ctk.CTkLabel(self.manual_scroll_frame, text=text, font=("Arial", 12, "bold"), text_color=self.BRAND_COLORS["text"])
            lbl.grid(row=0, column=col, padx=10, pady=5, sticky="w")
            self.manual_header_widgets.append(lbl)
        
        self.manual_rows = [] # To store row widget dictionaries
        self.row_counter = 1

    def update_table_columns(self):
        """Update table columns - simplified for labels"""
        # For manual entry, we only need fields for labels
        columns = ('poNumber', 'productCode', 'quantity')
        
        self.manual_table['columns'] = columns
        self.manual_table['show'] = 'headings'
        
        # Set column headings and widths
        self.manual_table.heading('poNumber', text='Invoice No.')
        self.manual_table.heading('productCode', text='Product Code')
        self.manual_table.heading('quantity', text='Quantity')
        
        for col in columns:
            self.manual_table.column(col, width=120, minwidth=80)

    def on_mode_change(self):
        """Handle InBound/OutBound mode change"""
        # Reset data
        self.clear_all_pdfs()
        self.extracted_data = []
        self.lbl_status.configure(text="No files selected")
        self.txt_preview.delete("0.0", "end")
        self.btn_save_excel.configure(state="disabled", text=f"Save {os.path.basename(self.get_active_excel())}")
        self.btn_download_txt.configure(state="disabled")
        
        # Clear manual rows
        self.clear_all_rows()
        
        mode = self.doc_type.get()
        
        # Update manual table headers if they exist
        if hasattr(self, 'manual_header_widgets') and len(self.manual_header_widgets) > 1:
            inv_label = "Invoice No." if mode == "Invoice" else "DO Number"
            self.manual_header_widgets[1].configure(text=inv_label)

        # Show/hide entry mode frame - only for InBound
        if mode == "Invoice":
            self.entry_mode_frame.grid(row=1, column=0, padx=0, pady=10, sticky="ew")
            self.btn_download_labels.pack(side="left", padx=20, pady=20)
            self.on_entry_mode_change() # Ensure UI elements like sidebar are updated
        else:
            self.entry_mode_frame.grid_forget()
            self.entry_mode.set("PDF")  # Reset to PDF mode
            self.on_entry_mode_change()  # Update UI
            self.btn_download_labels.pack_forget()

    def on_entry_mode_change(self):
        """Handle PDF/Manual entry mode change"""
        mode = self.entry_mode.get()
        
        # Hide all content frames first
        self.upload_frame.grid_forget()
        self.manual_frame.grid_forget()
        self.email_frame.grid_forget()
        self.email_inputs_frame.grid_forget()
        self.txt_preview.grid_forget()
        self.manual_scroll_frame.grid_forget()
        self.file_sidebar.grid_forget()
        
        if mode == "PDF":
            # Show PDF upload UI
            self.upload_frame.grid(row=2, column=0, padx=0, pady=10, sticky="ew")
            self.preview_frame.grid(row=4, column=0, padx=0, pady=10, sticky="nsew")
            
            # Sidebar only for OutBound
            if self.doc_type.get() == "Packing List":
                self.file_sidebar.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
                self.txt_preview.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
                self.btn_upload.configure(text="Upload PDF(s)")
                self.btn_clear_pdfs.pack(side="left", padx=10, pady=20)
            else:
                self.file_sidebar.grid_forget()
                self.txt_preview.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
                self.btn_upload.configure(text="Upload PDF")
                self.btn_clear_pdfs.pack_forget()
            
            # Show all export buttons
            self.btn_save_excel.pack(side="left", padx=20, pady=20)
            self.btn_download_txt.pack(side="left", padx=20, pady=20)
            self.btn_download_labels.pack(side="left", padx=20, pady=20)
        elif mode == "Manual":
            # Show manual entry UI (labels only)
            self.manual_frame.grid(row=2, column=0, padx=0, pady=10, sticky="ew")
            self.preview_frame.grid(row=4, column=0, padx=0, pady=10, sticky="nsew")
            # Use columnspan=2 to ensure it takes full weighted width
            self.manual_scroll_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
            
            # Hide Save to Excel and Download TXT buttons
            self.btn_save_excel.pack_forget()
            self.btn_download_txt.pack_forget()
            
            # Enable Download Labels if there's data
            if len(self.manual_rows) > 0:
                self.btn_download_labels.configure(state="normal")
            else:
                self.btn_download_labels.configure(state="disabled")
        elif mode == "Email":
            # ... (rest of Email mode)
            self.email_frame.grid(row=2, column=0, padx=0, pady=10, sticky="ew")
            self.email_inputs_frame.grid(row=3, column=0, padx=0, pady=5, sticky="ew")
            # Clear preview/table for email mode
            self.preview_frame.grid_forget()
            
            # Hide export buttons
            self.btn_save_excel.pack_forget()
            self.btn_download_txt.pack_forget()
            self.btn_download_labels.pack_forget()

    def clear_all_pdfs(self):
        """Reset PDF data and sidebar"""
        self.pdf_paths = []
        self.extracted_data = []
        self.current_preview_file = None
        for btn in self.file_btns.values():
            btn.destroy()
        self.file_btns = {}
        self.txt_preview.delete("0.0", "end")
        self.lbl_status.configure(text="No files selected")
        self.btn_save_excel.configure(state="disabled")
        self.btn_download_txt.configure(state="disabled")
        self.btn_download_labels.configure(state="disabled")

    def add_manual_row(self):
        """Add a new empty row to manual entry with widgets"""
        row = self.row_counter
        
        # Product Code ComboBox
        p_code_cb = ctk.CTkComboBox(self.manual_scroll_frame, values=self.product_codes, width=180)
        p_code_cb.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        p_code_cb.set("")
        
        # Invoice No Entry
        inv_no_entry = ctk.CTkEntry(self.manual_scroll_frame, width=180)
        inv_no_entry.grid(row=row, column=1, padx=10, pady=5, sticky="w")
        
        # Total Qty Entry
        total_qty_entry = ctk.CTkEntry(self.manual_scroll_frame, width=120)
        total_qty_entry.grid(row=row, column=2, padx=10, pady=5, sticky="w")
        
        # Pallet Qty Entry
        pallet_qty_entry = ctk.CTkEntry(self.manual_scroll_frame, width=120)
        pallet_qty_entry.grid(row=row, column=3, padx=10, pady=5, sticky="w")
        
        # Delete Button
        delete_btn = ctk.CTkButton(self.manual_scroll_frame, text="X", width=40, font=("Arial", 12, "bold"), 
                                   fg_color="red", hover_color="#8b0000")
        delete_btn.grid(row=row, column=4, padx=5, pady=5)
        
        # Store widgets
        row_data = {
            'productCode': p_code_cb,
            'poNumber': inv_no_entry,
            'quantity': total_qty_entry,
            'palletQuantity': pallet_qty_entry,
            'delete_btn': delete_btn
        }
        
        # Bindings
        p_code_cb.configure(command=lambda val, r=row_data: self.on_manual_product_select(val, r))
        # Add filtering logic for the dropdown
        p_code_cb.bind('<KeyRelease>', lambda e, c=p_code_cb: self.filter_product_codes(e, c))
        
        delete_btn.configure(command=lambda r=row_data: self.delete_manual_row(r))
        
        self.manual_rows.append(row_data)
        self.row_counter += 1
        
        # Enable download labels button
        self.btn_download_labels.configure(state="normal")
        
        return row_data

    def on_manual_product_select(self, selected_code, row_data):
        """Auto-fill pallet quantity and description when product is selected"""
        if selected_code in self.product_qty_mapping:
            mapping_info = self.product_qty_mapping[selected_code]
            qty = mapping_info['qty']
            row_data['palletQuantity'].delete(0, 'end')
            row_data['palletQuantity'].insert(0, str(qty))
        
        # Reset dropdown values to full list after selection
        row_data['productCode'].configure(values=self.product_codes)

    def filter_product_codes(self, event, cb):
        """Filter the product code dropdown values based on user input"""
        val = cb.get()
        if not val:
            cb.configure(values=self.product_codes)
        else:
            filtered = [c for c in self.product_codes if val.lower() in c.lower()]
            # Only update if the list has changed to avoid flickering
            if filtered != list(cb.cget("values")):
                cb.configure(values=filtered)
        
        # Note: CTkComboBox doesn't automatically open the dropdown on typing.
        # The user can click the dropdown arrow to see filtered results.

    def delete_manual_row(self, row_data):
        """Delete a single row of widgets"""
        # Destroy all widgets in that row
        for widget in row_data.values():
            widget.destroy()
        
        # Remove from tracking list
        if row_data in self.manual_rows:
            self.manual_rows.remove(row_data)
        
        # Disable download labels if no rows left
        if len(self.manual_rows) == 0:
            self.btn_download_labels.configure(state="disabled")

    def delete_selected_row(self):
        """Dummy method to satisfy old binding if any - now handled per-row"""
        pass

    def clear_all_rows(self):
        """Clear all rows from manual entry"""
        for row_data in self.manual_rows[:]: # Copy list to iterate
            self.delete_manual_row(row_data)
        
        self.btn_download_labels.configure(state="disabled")
        self.row_counter = 1

    def on_cell_double_click(self, event):
        """Deprecated for widget-based manual entry"""
        pass

    def validate_date(self, date_str):
        """Validate date format (DD-MMM-YYYY) and ensure it's not in the future"""
        try:
            # Try parsing with the desired format
            input_date = datetime.strptime(date_str, "%d-%b-%Y")
            # Clear time for comparison
            today = datetime.now().replace(hour=23, minute=59, second=59, microsecond=999999)
            if input_date > today:
                return False, "GRN Date cannot be in the future."
            return True, input_date.strftime("%d-%b-%Y")
        except ValueError:
            return False, "Invalid date format. Use DD-MMM-YYYY (e.g., 02-Feb-2026)."

    def get_active_excel(self):
        return self.excel_invoice if self.doc_type.get() == "Invoice" else self.excel_packing

    def upload_pdf(self):
        if self.doc_type.get() == "Packing List":
            file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
            if file_paths:
                self.process_multiple_files(file_paths)
        else:
            # Single file for InBound
            file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if file_path:
                self.clear_all_pdfs() # Ensure only one file's data exists
                self.process_multiple_files([file_path])
                self.lbl_status.configure(text=f"Selected: {os.path.basename(file_path)}")

    def process_multiple_files(self, file_paths):
        loading_popup = messagebox.showinfo("Processing", f"Processing {len(file_paths)} files. Please wait...")
        
        mode = self.doc_type.get()
        newly_processed = 0
        
        for path in file_paths:
            if path in self.pdf_paths:
                continue
            
            try:
                data = extract_all_data(path, doc_type=mode)
                if data:
                    self.pdf_paths.append(path)
                    self.extracted_data.extend(data)
                    filename = os.path.basename(path)
                    
                    # Add sidebar button
                    btn = ctk.CTkButton(self.file_sidebar, text=filename, fg_color="transparent", 
                                        text_color=self.BRAND_COLORS["text"],
                                        anchor="w", command=lambda p=path: self.switch_preview(p))
                    btn.pack(fill="x", padx=5, pady=2)
                    self.file_btns[path] = btn
                    
                    if not self.current_preview_file:
                        self.current_preview_file = path
                        self.switch_preview(path)
                    
                    newly_processed += 1
            except Exception as e:
                print(f"Error processing {path}: {e}")
                
        if newly_processed > 0:
            self.btn_save_excel.configure(state="normal")
            self.btn_download_txt.configure(state="normal")
            if mode == "Invoice":
                self.btn_download_labels.configure(state="normal")
            self.lbl_status.configure(text=f"Total: {len(self.pdf_paths)} files, {len(self.extracted_data)} products")
        else:
            messagebox.showwarning("Warning", "No new products found in the selected files.")

    def switch_preview(self, file_path):
        """Switch the current textbox preview to the selected file"""
        # Save current edits before switching
        self.sync_from_preview()
        
        # Highlight selected button
        for path, btn in self.file_btns.items():
            if path == file_path:
                btn.configure(fg_color=self.BRAND_COLORS["primary"], text_color=self.BRAND_COLORS["white"]) # Brand blue for selected
            else:
                btn.configure(fg_color="transparent", text_color=self.BRAND_COLORS["text"])
        
        self.current_preview_file = file_path
        self.update_preview()

    def upload_email_pdfs(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if file_paths:
            # Append uniquely
            for path in file_paths:
                if path not in self.email_pdf_paths:
                    self.email_pdf_paths.append(path)
                    
                    # Try to extract info from the FIRST new file if fields are currently empty
                    if not self.email_supplier_entry.get() or not self.email_invoice_entry.get():
                        filename = os.path.basename(path)
                        name_no_ext = os.path.splitext(filename)[0].lower()
                        
                        # ONLY extract if it's NOT a receipt or slip
                        if not re.search(r'(receipt|slip)', name_no_ext):
                            # Search for ALL invoice numbers (e.g., starts with letters, ends with digits, min 5 digits total)
                            inv_matches = re.findall(r'([a-z]+[0-9]{5,}|[0-9]{7,})', name_no_ext)
                            
                            if inv_matches:
                                 # Join all unique matches
                                 unique_invs = []
                                 for m in inv_matches:
                                     m_upper = m.upper()
                                     if m_upper not in unique_invs:
                                         unique_invs.append(m_upper)
                                 
                                 extracted_inv = " & ".join(unique_invs)
                                 if not self.email_invoice_entry.get():
                                     self.email_invoice_entry.insert(0, extracted_inv)
                                 
                                 if not self.email_supplier_entry.get():
                                     # Take everything before the FIRST invoice number as supplier
                                     name_orig = os.path.splitext(os.path.basename(path))[0]
                                     first_match_lower = inv_matches[0].lower()
                                     match_start = name_orig.lower().find(first_match_lower)
                                     supplier_raw = name_orig[:match_start].strip(" -_")
                                     
                                     # Clean up common unwanted words
                                     unwanted_words = ["INVOICE NO", "INVOICE", "INV NO", "INV"]
                                     supplier_clean = supplier_raw
                                     for word in unwanted_words:
                                         supplier_clean = re.sub(rf'(?i){word}', '', supplier_clean).strip(" -_")
                                     
                                     if supplier_clean:
                                         self.email_supplier_entry.insert(0, supplier_clean.title())
                        # If no match found, we leave them empty as requested

            self.lbl_email_status.configure(text=f"Ready: {len(self.email_pdf_paths)} files selected")

    def clear_email_files(self):
        self.email_pdf_paths = []
        self.email_supplier_entry.delete(0, 'end')
        self.email_invoice_entry.delete(0, 'end')
        self.lbl_email_status.configure(text="No files selected")

    def create_outlook_draft(self):
        if not self.email_pdf_paths:
            messagebox.showwarning("Warning", "Please select PDF files first.")
            return

        supplier = self.email_supplier_entry.get().strip() or "[Supplier Name]"
        invoice = self.email_invoice_entry.get().strip() or "[Invoice No]"

        if win32com is None:
            messagebox.showerror("Error", "Outlook integration (pywin32) is not available. Please install it.")
            return

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0) # 0 = olMailItem
            
            # Check if any filename contains 'receipt' or 'slip' to adjust subject and body
            has_receipt = any(re.search(r'(receipt|slip)', os.path.basename(p).lower()) for p in self.email_pdf_paths)
            receipt_suffix = " & RECEIPT SLIP" if has_receipt else ""
            
            # Determine if we say 'Invoice No' or 'Invoice Nos'
            inv_label = "INVOICE NOS" if "&" in invoice or "," in invoice else "INVOICE NO"
            mail.Subject = f"NEW MATERIAL INWARD FROM {supplier.upper()} {inv_label}-{invoice}{receipt_suffix}"
            
            # Add Recipients
            mail.To = "receiving@advics-ind.co.in; khageshkant_sharma@advics-ind.co.in"
            mail.CC = "ranjeet_sandhu@advics-ind.co.in; admin.gj1@nagarkot.co.in; logistic_planning@advics-ind.co.in"
            # Adjust body based on receipt slip presence
            receipt_body_part = " and the corresponding receipt slips" if has_receipt else ""
            
            # Reverted to simple body content, Outlook will append your default signature automatically
            inv_body_label = "Invoice Nos" if "&" in invoice or "," in invoice else "Invoice No"
            body_content = f"""
            <p>Dear San,</p>
            <p>Greetings!!</p>
            <p>Please find attached the confirmation of material received from {supplier}, along with {inv_body_label}: {invoice}{receipt_body_part}.</p>
            """
            
            # Attach files
            for path in self.email_pdf_paths:
                if os.path.exists(path):
                    mail.Attachments.Add(path)
                else:
                    print(f"File not found: {path}")

            # This order (Display then HTMLBody) is key to preserving default Outlook signatures 
            # while adding our custom text on top
            mail.Display()
            mail.HTMLBody = body_content + mail.HTMLBody
            try:
                # This helps bring the specific window to the front
                mail.GetInspector.Activate()
            except Exception:
                pass
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create Outlook draft:\n{e}")

    def process_file(self):
        """Deprecated: Use process_multiple_files instead"""
        pass

    def update_preview(self):
        self.txt_preview.delete("0.0", "end")
        if not self.extracted_data or not self.current_preview_file:
            return
            
        filename = os.path.basename(self.current_preview_file)
        # Filter master list for current file
        file_data = [item for item in self.extracted_data if item.get('fileName') == filename]
        
        if not file_data:
            return

        df = pd.DataFrame(file_data)
        # Remove fileName from preview to save space
        if 'fileName' in df.columns:
            df = df.drop(columns=['fileName'])
            
        df = df.fillna('-').replace(r'^\s*$', '-', regex=True)
        table_str = df.to_string(index=False, justify='left', col_space=15)
        self.txt_preview.insert("0.0", table_str)

    def get_manual_data(self):
        """Extract and validate data from manual entry widgets"""
        data = []
        for row_data in self.manual_rows:
            p_code = row_data['productCode'].get().strip()
            inv_no = row_data['poNumber'].get().strip()
            total_qty_str = row_data['quantity'].get().strip()
            pallet_qty_str = row_data['palletQuantity'].get().strip()
            
            # Mandatory check
            if not p_code or not inv_no or not total_qty_str or not pallet_qty_str:
                messagebox.showwarning("Incomplete Data", "All fields (Product Code, Invoice No, Total Qty, Pallet Qty) are mandatory for all rows.")
                return None
            
            try:
                total_qty = float(total_qty_str)
                pallet_qty = float(pallet_qty_str)
            except ValueError:
                messagebox.showerror("Invalid Quantity", f"Invalid quantity in row with invoice {inv_no}. Please enter numbers.")
                return None
                
            # Perform splitting logic similar to PDF mode if needed
            # For simplicity, we create a list of items where each has 'quantity' as current_qty
            # and we replicate based on splitting logic in download_labels
            
            row_dict = {
                'productCode': p_code,
                'total_quantity': total_qty,
                'pallet_quantity': pallet_qty,
                'quantity': total_qty_str # Keep original total for reference
            }
            
            if self.doc_type.get() == "Invoice":
                row_dict['poNumber'] = inv_no
            else:
                row_dict['doNumber'] = inv_no
                
            data.append(row_dict)
        
        return data

    def save_excel(self):
        mode = self.doc_type.get()
        target = self.get_active_excel()
        
        # Get data based on entry mode
        if self.entry_mode.get() == "Manual":
            data = self.get_manual_data()
        else:
            self.sync_from_preview()
            data = self.extracted_data
        
        if update_excel(target, data, doc_type=mode):
            messagebox.showinfo("Success", f"Data successfully added to {os.path.basename(target)}")
        else:
            messagebox.showerror("Error", "Failed to update Excel file.")

    def download_txt(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if save_path:
            try:
                if self.entry_mode.get() == "Manual":
                    data = self.get_manual_data()
                else:
                    # Sync current preview before downloading all
                    self.sync_from_preview()
                    data = self.extracted_data
                
                df = pd.DataFrame(data)
                
                # Filter columns to match active mode
                if self.doc_type.get() == "Invoice":
                    cols = ['storerCode', 'warehouseCode', 'poNumber', 'poDate', 'supplierCode', 'otherReference', 'productCode', 'quantity', 'uomCode']
                else:
                    cols = ['storerCode', 'warehouseCode', 'doNumber', 'consigneeCode', 'shipToAddressPostCode', 'requiredDate', 'otherReference', 'productCode', 'quantity', 'uomCode']
                
                valid_cols = [c for c in cols if c in df.columns]
                df = df[valid_cols]
                
                df.to_csv(save_path, sep='\t', index=False)
                messagebox.showinfo("Success", f"File saved to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

    def sync_from_preview(self):
        """Attempts to parse the current text preview back into extracted_data for the specific file."""
        try:
            if not self.current_preview_file:
                return
                
            content = self.txt_preview.get("0.0", "end").strip()
            if not content:
                return
            
            # Parse fixed-width/space-separated table
            df = pd.read_csv(io.StringIO(content), sep=r'\s+', dtype=str)
            
            # Add back the filename
            filename = os.path.basename(self.current_preview_file)
            df['fileName'] = filename
            
            # Convert back to list of dicts
            updated_file_data = df.to_dict('records')
            
            # Merge back into master list: remove old items for this file and add updated ones
            self.extracted_data = [item for item in self.extracted_data if item.get('fileName') != filename]
            self.extracted_data.extend(updated_file_data)
            
            # Update status
            self.lbl_status.configure(text=f"Total: {len(self.pdf_paths)} files, {len(self.extracted_data)} products (Synced)")
        except Exception as e:
            print(f"Sync error: {e}")
            pass

    def download_labels(self):
        """Generate labels from either PDF data or manual entry"""
        # Check if we're in manual entry mode
        if self.entry_mode.get() == "Manual":
            # Get and validate GRN Date
            grn_date_val = self.grn_date_manual.get().strip()
            is_valid, grn_date = self.validate_date(grn_date_val)
            if not is_valid:
                messagebox.showerror("Invalid Date", grn_date)
                return
            
            # Get data from manual table
            manual_data_raw = self.get_manual_data()
            if manual_data_raw is None: # Validation failed
                return
            if not manual_data_raw:
                messagebox.showwarning("Warning", "No data to generate labels. Please add rows first.")
                return
            
            # Apply splitting logic to manual data
            final_manual_data = []
            for item in manual_data_raw:
                total_qty = item['total_quantity']
                pallet_qty = item['pallet_quantity']
                
                if pallet_qty <= 0:
                    updated_item = item.copy()
                    final_manual_data.append(updated_item)
                    continue
                
                remaining = total_qty
                while remaining > 0:
                    chunk = min(remaining, pallet_qty)
                    updated_item = item.copy()
                    updated_item['quantity'] = str(int(chunk)) if chunk.is_integer() else str(chunk)
                    final_manual_data.append(updated_item)
                    remaining -= chunk
                    if remaining < 0.001: remaining = 0
            
            # Ask for save path
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not save_path:
                return
            
            # Generate labels
            for item in final_manual_data:
                # Add description and total from mapping
                p_code = item.get('productCode')
                if p_code in self.product_qty_mapping:
                    item['description'] = self.product_qty_mapping[p_code]['desc']
                else:
                    item['description'] = ""
                
                # In manual mode, total_quantity is stored as 'total_quantity' in item
                # ensure it's a string for label_utils
                tq = item.get('total_quantity', 0)
                item['total_quantity'] = str(int(tq)) if isinstance(tq, (int, float)) and float(tq).is_integer() else str(tq)

            if generate_label_pdf(final_manual_data, save_path, grn_date=grn_date):
                messagebox.showinfo("Success", f"Labels saved to {save_path}\nGenerated {len(final_manual_data)} labels.")
            else:
                messagebox.showerror("Error", "Failed to generate labels.")
        else:
            # Original PDF mode with review popup
            # Sync changes from text preview first
            self.sync_from_preview()
            
            if not self.extracted_data:
                return
                
            # Create a Review Popup
            self.review_window = ctk.CTkToplevel(self)
            self.review_window.title("Review Quantities")
            self.review_window.geometry("800x600")
            
            # Make the popup modal and bring to front
            self.review_window.lift()
            self.review_window.focus_force()
            self.review_window.grab_set()
            
            # Load excel mapping once
            qty_mapping = load_quantity_mapping()
            
            ctk.CTkLabel(self.review_window, text="Verify Quantities before Generating Labels", font=("Arial", 16, "bold")).pack(pady=10)
            
            # GRN Date in popup
            date_frame = ctk.CTkFrame(self.review_window, fg_color="transparent")
            date_frame.pack(pady=5)
            ctk.CTkLabel(date_frame, text="GRN Date:").pack(side="left", padx=5)
            self.pdf_grn_date_entry = ctk.CTkEntry(date_frame, width=120)
            self.pdf_grn_date_entry.pack(side="left", padx=5)
            self.pdf_grn_date_entry.insert(0, datetime.now().strftime("%d-%b-%Y"))
            
            # Scrollable list
            scroll_frame = ctk.CTkScrollableFrame(self.review_window)
            scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Grid Configuration
            scroll_frame.grid_columnconfigure(0, weight=1) # Index
            scroll_frame.grid_columnconfigure(1, weight=3) # Product Code
            scroll_frame.grid_columnconfigure(2, weight=3) # Invoice No
            scroll_frame.grid_columnconfigure(3, weight=2) # Total Qty
            # Column 4 and onwards (Dynamic Palette Qty) should not expand excessively
            # giving them small or equal weight keeps them packed to the left
            for i in range(4, 20):
                scroll_frame.grid_columnconfigure(i, weight=0) 
            
            # Headers
            inv_do_label = "Invoice No." if self.doc_type.get() == "Invoice" else "DO Number"
            headers = ["#", "Product Code", inv_do_label, "Total Quantity", "Pallet Quantity"]
            for col, text in enumerate(headers):
                ctk.CTkLabel(scroll_frame, text=text, font=("Arial", 12, "bold")).grid(row=0, column=col, padx=5, pady=5, sticky="w")
            
            self.qty_entries = [] # To store (item_dict, [entry_widgets])
            
            # Helper to create row labels
            def create_row_labels(idx, item_data, display_qty):
                row = idx + 1
                p_c = str(item_data.get('productCode', 'Unknown'))
                i_n = str(item_data.get('poNumber', item_data.get('doNumber', 'Unknown')))
                
                ctk.CTkLabel(scroll_frame, text=str(idx+1)).grid(row=row, column=0, padx=5, pady=2, sticky="w")
                ctk.CTkLabel(scroll_frame, text=p_c).grid(row=row, column=1, padx=5, pady=2, sticky="w")
                ctk.CTkLabel(scroll_frame, text=i_n).grid(row=row, column=2, padx=5, pady=2, sticky="w")
                ctk.CTkLabel(scroll_frame, text=str(display_qty)).grid(row=row, column=3, padx=5, pady=2, sticky="w")
                
            for i, item in enumerate(self.extracted_data):
                p_code = str(item.get('productCode', 'Unknown'))
                total_qty_str = str(item.get('quantity', '0'))
                try:
                    total_qty = float(total_qty_str)
                except:
                    total_qty = 0
                
                mapped_val = qty_mapping.get(p_code)
                
                is_special_split = (p_code == '116040-92240')
                entries_for_row = []
                
                create_row_labels(i, item, total_qty_str)
                row_idx = i + 1

                if is_special_split and mapped_val:
                    try:
                        pallet_qty = float(mapped_val)
                        if pallet_qty > 0:
                            # Split logic to determine HOW MANY entries
                            remaining = total_qty
                            col_offset = 4
                            
                            # If total is 0, at least show one empty/zero box? 
                            # Assuming total > 0 usually.
                            while remaining > 0:
                                chunk = min(remaining, pallet_qty)
                                chunk_str = str(int(chunk)) if chunk.is_integer() else str(chunk)
                                
                                qty_entry = ctk.CTkEntry(scroll_frame, width=100)
                                qty_entry.grid(row=row_idx, column=col_offset, padx=2, pady=2, sticky="w")
                                qty_entry.insert(0, chunk_str)
                                entries_for_row.append(qty_entry)
                                
                                col_offset += 1
                                remaining -= chunk
                                if remaining < 0.001: remaining = 0
                    except ValueError:
                        # Fallback if mapped_val parsing fails
                        default_pallet_qty = total_qty_str
                        qty_entry = ctk.CTkEntry(scroll_frame, width=100)
                        qty_entry.grid(row=row_idx, column=4, padx=5, pady=2, sticky="w")
                        qty_entry.insert(0, str(default_pallet_qty))
                        entries_for_row.append(qty_entry)
                
                else:
                    # Standard single entry
                    default_entry_val = mapped_val if mapped_val else total_qty_str
                    qty_entry = ctk.CTkEntry(scroll_frame, width=100)
                    qty_entry.grid(row=row_idx, column=4, padx=5, pady=2, sticky="w")
                    qty_entry.insert(0, str(default_entry_val))
                    entries_for_row.append(qty_entry)
                
                # If for some reason special logic didn't add any entries (e.g. pallet=0 or total=0), add one default
                if not entries_for_row:
                     qty_entry = ctk.CTkEntry(scroll_frame, width=100)
                     qty_entry.grid(row=row_idx, column=4, padx=5, pady=2, sticky="w")
                     qty_entry.insert(0, "0")
                     entries_for_row.append(qty_entry)

                self.qty_entries.append((item, entries_for_row))
                
            # Confirm Button
            btn_confirm = ctk.CTkButton(self.review_window, text="Confirm & Generate PDF", command=self.generate_updated_pdf, fg_color="green")
            btn_confirm.pack(pady=20)
            
    def generate_updated_pdf(self):
        # Validate GRN Date from popup
        grn_date_val = self.pdf_grn_date_entry.get().strip()
        is_valid, grn_date = self.validate_date(grn_date_val)
        if not is_valid:
            messagebox.showerror("Invalid Date", grn_date)
            return
            
        # 1. Collect updated data with automatic splitting
        final_data = []
        
        for item, entry_list in self.qty_entries:
            # Get the pallet quantity from the FIRST entry (user can edit this)
            if not entry_list:
                continue
                
            pallet_qty_str = entry_list[0].get().strip()
            try:
                pallet_qty = float(pallet_qty_str)
            except ValueError:
                pallet_qty = 0
            
            # Get the TOTAL quantity from the original item data
            try:
                total_qty = float(item.get('quantity', '0'))
            except:
                total_qty = 0
            
            # If pallet qty is invalid or 0, just create one label with total
            if pallet_qty <= 0:
                if total_qty > 0:
                    updated_item = item.copy()
                    updated_item['quantity'] = str(int(total_qty)) if total_qty.is_integer() else str(total_qty)
                    final_data.append(updated_item)
                continue
            
            # Split logic: Create labels based on Total / Pallet
            remaining = total_qty
            max_labels = 1000  # Safety limit
            count = 0
            
            while remaining > 0 and count < max_labels:
                # Determine quantity for this specific label
                current_label_qty = min(remaining, pallet_qty)
                
                # Create label item
                updated_item = item.copy()
                qty_str = str(int(current_label_qty)) if current_label_qty.is_integer() else str(current_label_qty)
                updated_item['quantity'] = qty_str
                updated_item['total_quantity'] = str(int(total_qty)) if total_qty.is_integer() else str(total_qty)
                
                # Add description
                p_code = str(item.get('productCode'))
                if p_code in self.product_qty_mapping:
                    updated_item['description'] = self.product_qty_mapping[p_code]['desc']
                else:
                    updated_item['description'] = ""
                    
                final_data.append(updated_item)
                
                remaining -= current_label_qty
                count += 1
                
                # Precision safety for floats
                if remaining < 0.001:
                    remaining = 0

        # 2. Ask for save path
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not save_path:
            return # User cancelled file dialog
            
        # 3. Generate
        if generate_label_pdf(final_data, save_path, grn_date=grn_date):
            self.review_window.destroy()
            messagebox.showinfo("Success", f"Labels saved to {save_path}\nGenerated {len(final_data)} labels.")
        else:
            messagebox.showerror("Error", "Failed to generate labels.")

if __name__ == "__main__":
    app = ASNGeneratorApp()
    app.mainloop()
