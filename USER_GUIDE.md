# InBound & OutBound Data Generator User Guide

## Introduction
The **InBound & OutBound Data Generator** is a specialized tool designed for Nagarkot Forwarders Pvt Ltd to streamline the extraction of ASN (Advanced Shipping Notice) data from PDF invoices and packing lists. It automates the process of updating Excel registers and generating pallet labels for warehouse operations.

## How to Use

### 1. Launching the App
Download and run the provided `asn_gui_v2.exe`. The application will open in full-screen mode with the Nagarkot corporate branding.

### 2. The Workflow (Step-by-Step)

#### A. InBound Data Processing (Invoices)
1. **Select Mode**: Click the **InBound** radio button.
2. **Choose Entry Mode**:
   - **PDF Upload**: For automatic extraction from PDF invoices.
   - **Manual Entry**: For generating labels manually without a PDF.
   - **Email Draft**: To quickly generate an Outlook draft with attachments.
3. **Action**: 
   - If **PDF Upload**: Click **Upload PDF**, select your invoice. Review the data in the preview box.
   - If **Manual Entry**: Click **Add Row**, select a **Product Code** (this auto-fills pallet quantity), enter the **Invoice No** and **Total Quantity**.
4. **Export**: Click **Save to Excel** to update the `InBound ASN Register.xlsx` or **Download Labels** to generate a PDF of pallet labels.

#### B. OutBound Data Processing (Packing Lists)
1. **Select Mode**: Click the **OutBound** radio button.
2. **Action**: Click **Upload PDF(s)**. You can select multiple files at once.
3. **Review**: Use the sidebar on the left to switch between different loaded documents. You can edit the text in the preview box if corrections are needed; changes are synced automatically when switching files or saving.
4. **Export**: Click **Save to Excel** to update `OutBound Order Register.xlsx` or **Download (Tab Delimited)** for system uploads.

#### C. Creating Outlook Drafts
1. **Switch to Email Draft**: Select **InBound** then **Email Draft**.
2. **Upload Files**: Click **Select PDF Files** to add invoices and receipt slips.
3. **Extraction**: The tool will attempt to extract the Supplier Name and Invoice Number from the filenames.
4. **Create**: Click **Create Outlook Draft** to open a pre-composed email in Microsoft Outlook with all attachments and correct recipients.

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **InBound / OutBound** | Switches between receiving and dispatching workflows. | Choice |
| **PDF / Manual / Email** | Switches between automation modes. | Choice |
| **Upload PDF(s)** | Opens file dialog to select input documents. | `.pdf` |
| **Product Code** | Searchable dropdown for item selection. | `6-5 Digit` format |
| **Total Quantity** | Total count of items for the specific invoice row. | Number |
| **Pallet Quantity** | Number of items per pallet (auto-filled from mapping). | Number |
| **GRN Date** | Date of material receipt. | `DD-MMM-YYYY` |

## Troubleshooting & Validations

If you see an error, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| `Invalid date format` | The GRN Date entered doesn't match the format. | Use `02-Feb-2026` style. |
| `GRN Date cannot be in the future` | You've entered a date after today. | Ensure the date is today or earlier. |
| `No new products found` | The PDF could not be parsed or contains no matching codes. | Verify the PDF is a supported format. |
| `All fields are mandatory` | A row in Manual Entry is missing data. | Ensure all columns in the row are filled. |
| `Outlook integration... not available` | `pywin32` is missing or Outlook is not installed. | Ensure Microsoft Outlook is installed on the PC. |

## Contact Support
For feedback or technical issues, please contact the Nagarkot IT team.
