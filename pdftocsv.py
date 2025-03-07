import os
import re
import csv
import PyPDF2

# Directory containing PDFs (Update this if needed)
pdf_folder = r"D:\Code\Invoices"
output_csv = "invoices_data.csv"

def extract_text_from_pdf(pdf_path):
    """
    Extracts text from a PDF file.

    Parameters:
        pdf_path (str): Path to the PDF file.

    Returns:
        str: Extracted text from the PDF.
    """
    text = ""
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            extracted_text = page.extract_text()
            if extracted_text:
                text += extracted_text + "\n"
    return text

def extract_invoice_details(text):
    """
    Extracts invoice details using regex patterns.

    Parameters:
        text (str): Extracted text from the invoice PDF.

    Returns:
        dict: Dictionary containing extracted invoice details.
    """
    invoice_number = re.search(r"Invoice Number:\s*(.+)", text)
    due_date = re.search(r"Due Date:\s*(\d{2}/\d{2}/\d{4})", text)

    # Extract only the company name from "Bill To"
    bill_to_match = re.search(r"Bill To:\s*\n(\S+)", text)
    bill_to = bill_to_match.group(1).strip() if bill_to_match else ""

    # Extract PO Number (from Software Development services)
    po_number_match = re.search(r"Software Development services:\s*\n?([A-Za-z0-9\-/\.]+)", text)   ##re.search(r"Software Development services:\s*([\w\-/\.]+)", text)
    po_number = po_number_match.group(1).strip() if po_number_match else ""

    # Extract total amount due
    total_amount_match = re.search(r"Total Amount Due\s*\$([\d,]+\.\d{2})", text)
    total_amount = total_amount_match.group(1).strip() if total_amount_match else ""

    return {
        "Invoice Number": invoice_number.group(1).strip() if invoice_number else "",
        "Due Date": due_date.group(1).strip() if due_date else "",
        "Bill To": bill_to,  # ✅ Extracts only the company name
        "PO Number": po_number,  # ✅ Extracts only the reference number
        "Total Amount Due": total_amount
    }

# Process all PDFs and store data
data_list = []
for filename in os.listdir(pdf_folder):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, filename)
        text = extract_text_from_pdf(pdf_path)
        extracted_data = extract_invoice_details(text)
        data_list.append(extracted_data)

# Write to CSV file
with open(output_csv, mode="w", newline="") as file:
    writer = csv.DictWriter(file, fieldnames=["Invoice Number", "Due Date", "Bill To", "PO Number", "Total Amount Due"])
    writer.writeheader()
    writer.writerows(data_list)

print(f"✅ Data successfully extracted and saved to {output_csv}")