import os
import re
import csv
from docx import Document

# Directory containing Word documents (Update this if needed)
docx_folder = r"D:\Code\Invoices"
output_csv = r"D:\Code\InvoicesCSV\invoices_data.csv"

def extract_text_from_docx(docx_path):
    """
    Extracts text from a Word (.docx) document.

    Parameters:
        docx_path (str): Path to the .docx file.

    Returns:
        str: Extracted text from the document.
    """
    text = []
    doc = Document(docx_path)
    for para in doc.paragraphs:
        text.append(para.text)
    return "\n".join(text), doc

def extract_invoice_details(text, doc):
    """
    Extracts invoice details using regex and extracts PO Number from tables.

    Parameters:
        text (str): Extracted text from the Word document.
        doc (Document): The Word document object.

    Returns:
        dict: Dictionary containing extracted invoice details.
    """
    invoice_number_match = re.search(r"Invoice Number:\s*(.+)", text)
    due_date_match = re.search(r"Due Date:\s*(\d{2}/\d{2}/\d{4})", text)

    # Extract Bill To (Company Name)
    bill_to_match = re.search(r"Bill To:\s*\n(.+)", text)
    bill_to = bill_to_match.group(1).strip() if bill_to_match else ""

    # Extract PO Number from the first row of tables
    po_number = "NOT FOUND"
    for table in doc.tables:
        if len(table.rows) > 1:
            first_row_desc = table.rows[1].cells[0].text.strip()
            if ":" in first_row_desc:
                po_number = first_row_desc.split(":", 1)[1].strip()
                break  # Stop after finding the first match

    # Extract Total Amount Due
    total_amount_match = re.search(r"Total Amount Due\s*\$([\d,]+\.\d{2})", text)
    total_amount = total_amount_match.group(1).strip() if total_amount_match else ""

    return {
        "Invoice Number": invoice_number_match.group(1).strip() if invoice_number_match else "",
        "Due Date": due_date_match.group(1).strip() if due_date_match else "",
        "Bill To": bill_to,
        "PO Number": po_number,
        "Total Amount Due": total_amount
    }

# Process all .docx files and store data
data_list = []
for filename in os.listdir(docx_folder):
    if filename.endswith(".docx"):
        docx_path = os.path.join(docx_folder, filename)
        try:
            text, doc = extract_text_from_docx(docx_path)
            extracted_data = extract_invoice_details(text, doc)
            data_list.append(extracted_data)
        except Exception as e:
            print(f"❌ Error processing {filename}: {e}")

# Ensure output directory exists
os.makedirs(os.path.dirname(output_csv), exist_ok=True)

# Write to CSV file
with open(output_csv, mode="w", newline="", encoding="utf-8") as file:
    writer = csv.DictWriter(file, fieldnames=["Invoice Number", "Due Date", "Bill To", "PO Number", "Total Amount Due"])
    writer.writeheader()
    writer.writerows(data_list)

print(f"✅ Data successfully extracted and saved to {output_csv}")