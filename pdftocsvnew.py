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
        tuple: Extracted text and the Document object.
    """
    text = []
    doc = Document(docx_path)
    
    # Extract text from paragraphs
    for para in doc.paragraphs:
        text.append(para.text.strip())

    return "\n".join(text), doc

def extract_invoice_details(text, doc):
    """
    Extracts invoice details using regex and retrieves PO Number from tables.

    Parameters:
        text (str): Extracted text from the Word document.
        doc (Document): The Word document object.

    Returns:
        dict: Dictionary containing extracted invoice details.
    """
    invoice_number_match = re.search(r"Invoice Number:\s*(.+)", text)
    due_date_match = re.search(r"Due Date:\s*(\d{2}/\d{2}/\d{4})", text)

    # Extract "Bill To" company name
    bill_to_match = re.search(r"Bill To:\s*(.+)", text)
    bill_to = bill_to_match.group(1).strip() if bill_to_match else "NOT FOUND"

    # Extract PO Number from tables
    po_number = "NOT FOUND"
    for table in doc.tables:
        for row in table.rows:
            row_text = row.cells[0].text.strip()
            print(f"Checking table row: {row_text}")  # Debugging line
            if ":" in row_text:
                po_number = row_text.split(":", 1)[1].strip()
                break  # Stop after finding the first match

    # Extract Total Amount Due
    total_amount_match = re.search(r"Total Amount Due\s*\$([\d,]+\.\d{2})", text)
    total_amount = total_amount_match.group(1).strip() if total_amount_match else "NOT FOUND"

    extracted_data = {
        "Invoice Number": invoice_number_match.group(1).strip() if invoice_number_match else "NOT FOUND",
        "Due Date": due_date_match.group(1).strip() if due_date_match else "NOT FOUND",
        "Bill To": bill_to,
        "PO Number": po_number,
        "Total Amount Due": total_amount
    }

    print(f"Extracted Data: {extracted_data}")  # Debugging line
    return extracted_data

# Process all .docx files and store data
data_list = []
for filename in os.listdir(docx_folder):
    if filename.endswith(".docx"):
        docx_path = os.path.join(docx_folder, filename)
        try:
            print(f"\nProcessing: {filename}")  # Debugging line
            text, doc = extract_text_from_docx(docx_path)
            print(f"\nExtracted Text:\n{text}\n")  # Debugging line
            extracted_data = extract_invoice_details(text, doc)
            if any(value != "NOT FOUND" for value in extracted_data.values()):  # Ensure there's valid data
                data_list.append(extracted_data)
            else:
                print(f"⚠️ No valid data found in {filename}")
        except Exception as e:
            print(f"❌ Error processing {filename}: {e}")

# Ensure output directory exists
os.makedirs(os.path.dirname(output_csv), exist_ok=True)

# Write to CSV file only if data_list has values
if data_list:
    with open(output_csv, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["Invoice Number", "Due Date", "Bill To", "PO Number", "Total Amount Due"])
        writer.writeheader()
        writer.writerows(data_list)
    print(f"\n✅ Data successfully extracted and saved to {output_csv}")
else:
    print("\n⚠️ No valid data extracted. CSV file was not created.")