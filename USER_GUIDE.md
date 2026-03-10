# InBound & OutBound Data Generator User Guide

## Introduction
The **ASN_Label_OutBound DataGenerator** is a specialized tool designed for Nagarkot Forwarders Pvt Ltd to streamline the extraction of ASN (Advanced Shipping Notice) data from PDF invoices, packing lists, and warehouse receipt slips. It automates the process of updating Excel registers and generating pallet labels for warehouse operations.

## How to Use

### 1. Launching the App
Run the `ASN_Label_OutBound Data Generator.exe`. The application opens in a branded window with simple navigation.

### 2. The Workflow (Step-by-Step)

#### A. InBound Data Processing (Invoices & Receipt Slips)
1. **Select Mode**: Click the **InBound** radio button.
2. **Choose Entry Mode**:
   - **PDF Upload**: For automatic extraction from PDF documents.
   - **Manual Entry**: For generating labels manually without a PDF.
3. **Action (PDF Upload)**: 
   - Click **Upload PDF**.
   - **Standard Invoices**: The tool extracts all product data. You will see buttons to **Save to Excel** and **Download (Tab Delimited)**.
   - **Receipt Slips**: If the PDF contains "RECEIPT SLIP" in the header, the tool automatically switches to a **Label-Only Mode**.
     - It isolates the **"Receipt Details"** section.
     - It hides the Excel/TXT buttons to prevent unverified registration.
     - Click **Download Labels** to proceed to the review popup.
4. **Manual Entry**: Use this to quickly generate labels by selecting a **Product Code** (auto-fills pallet quantity) and entering the **Total Quantity**.

#### B. OutBound Data Processing (Packing Lists)
1. **Select Mode**: Click the **OutBound** radio button.
2. **Action**: Click **Upload PDF(s)** (Multiple selection supported).
3. **Export**: Click **Save to Excel** or **Download (Tab Delimited)**.

### 3. Smart Label Generation
When you click **Download Labels**:
1. A **Review Quantities** popup appears.
2. **Standard Pack Split**: The tool automatically calculates the number of labels required based on the **Pallet Quantity** (Standard Pack) from your `Label Quantity.xlsx` mapping.
3. **Manual Override**: You can adjust the quantities or the GRN date before hitting **Confirm & Generate PDF**.

## Interface Reference

| Control / Input | Description | Note |
| :--- | :--- | :--- |
| **InBound / OutBound** | Switches workspace mode. | |
| **PDF / Manual** | Automation vs. Manual entry. | |
| **Download Labels** | Generates PDF labels (10cm x 7.5cm). | Enabled for Receipt Slips & Manual Entry. |
| **Save to Excel** | Updates the local tracking register. | Hidden for Receipt Slips. |
| **GRN Date** | Receipt date printed on labels. | Default is today. |

## Troubleshooting & Validations

| Message / Issue | Reason | Solution |
| :--- | :--- | :--- |
| `Excel/TXT buttons hidden` | You uploaded a Receipt Slip. | This is intentional; Receipt Slips are for label generation only. |
| `Invalid Date` | GRN Date format is incorrect. | Use `DD-MMM-YYYY` (e.g., `10-Mar-2026`). |
| `No new products found` | PDF layout is unsupported. | Ensure the PDF follows the standard Invoice or Receipt Slip layout. |
| `Logo Load Error` | `logo.png` is missing. | Ensure the logo is in the same folder as the script/exe if running from source. |

## Contact Support
For feedback or technical issues, please contact the Nagarkot IT team.
