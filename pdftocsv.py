import os
import re
import csv
import pdfplumber

# Directory containing PDFs (Update this if needed)
pdf_folder = r"D:\Code\Invoices"

# ✅ Set the specific path where the CSV file should be saved
output_csv = r"D:\Code\InvoicesCSV\invoices_data.csv"

def extract_text_from_pdf(pdf_path):
    """
    Extracts text from a PDF file using pdfplumber.

    Parameters:
        pdf_path (str): Path to the PDF file.

    Returns:
        str: Extracted text from the PDF.
    """
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def extract_invoice_details(text):
    """
    Extracts invoice details using regex patterns.

    Parameters:
        text (str): Extracted text from the invoice PDF.

    Returns:
        dict: Dictionary containing extracted invoice details.
    """
    invoice_number_match = re.search(r"Invoice Number:\s*(.+)", text)
    due_date_match = re.search(r"Due Date:\s*(\d{2}/\d{2}/\d{4})", text)
    
    # Extract only the company name from "Bill To"
    bill_to_match = re.search(r"Bill To:\s*\n(\S+)", text)
    bill_to = bill_to_match.group(1).strip() if bill_to_match else ""

    # Extract PO Number from "Software Development services"
    po_number_match = re.search(r"PO Number[:\s]*([\w\-/\.]+)", text, re.IGNORECASE)
    if not po_number_match:
        po_number_match = re.search(r"Software Development services.*?([\w\-/\.]+)", text, re.IGNORECASE)  # Backup search
    po_number = po_number_match.group(1).strip() if po_number_match else "NOT FOUND"

    # Extract total amount due
    total_amount_match = re.search(r"Total Amount Due\s*\$([\d,]+\.\d{2})", text)
    total_amount = total_amount_match.group(1).strip() if total_amount_match else ""

    return {
        "Invoice Number": invoice_number_match.group(1).strip() if invoice_number_match else "",
        "Due Date": due_date_match.group(1).strip() if due_date_match else "",
        "Bill To": bill_to,
        "PO Number": po_number,
        "Total Amount Due": total_amount
    }

# Process all PDFs and store data
data_list = []
for filename in os.listdir(pdf_folder):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, filename)
        text = extract_text_from_pdf(pdf_path)

        # Debugging step: Print extracted text (Remove this after testing)
        print(f"\nExtracted Text from {filename}:\n{text}")

        extracted_data = extract_invoice_details(text)
        data_list.append(extracted_data)

# ✅ Ensure the output directory exists before saving the file
os.makedirs(os.path.dirname(output_csv), exist_ok=True)

# Write to CSV file
with open(output_csv, mode="w", newline="") as file:
    writer = csv.DictWriter(file, fieldnames=["Invoice Number", "Due Date", "Bill To", "PO Number", "Total Amount Due"])
    writer.writeheader()
    writer.writerows(data_list)

print(f"✅ Data successfully extracted and saved to {output_csv}")