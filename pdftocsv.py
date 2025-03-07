import os
import re
import csv
import PyPDF2

# Directory containing PDFs (Update this if needed)
pdf_folder = r"D:\Code\Invoices"
output_csv = "invoices_data.csv"

# Function to extract text from a PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text += page.extract_text() + "\n"
    return text

# Function to extract required details
def extract_invoice_details(text):
    invoice_number = re.search(r"Invoice Number:\s*(.+)", text)
    due_date = re.search(r"Due Date:\s*(\d{2}/\d{2}/\d{4})", text)
    bill_to = re.search(r"Bill To:\s*\n(\S.+)", text)       ##re.search(r"Bill To:\s*(.*?)\n\n", text, re.DOTALL)
    service = re.search(r"Software Development services:\s*([\w\-/\.]+)", text)     ##re.search(r"Software Development services:\s*([\w\-/]+)", text)
    total_amount = re.search(r"Total Amount Due\s*\$([\d,]+\.\d{2})", text)

    return {
        "Invoice Number": invoice_number.group(1) if invoice_number else "",
        "Due Date": due_date.group(1) if due_date else "",
        "Bill To": bill_to.group(1).replace("\n", ", ") if bill_to else "",
        "PO Number": service.group(1) if service else "",
        "Total Amount Due": total_amount.group(1) if total_amount else ""
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

print(f"Data successfully extracted and saved to {output_csv}")