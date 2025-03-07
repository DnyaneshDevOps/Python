import os
import re
import csv
import pdfplumber

# Directory containing PDFs
pdf_folder = r"D:\Code\Invoices"
output_csv = r"D:\Code\InvoicesCSV\invoices_data.csv"

def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF file using pdfplumber."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def extract_po_number_from_pdf(pdf_path):
    """Extracts PO Number from a table in the PDF (handles multi-line values)."""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()  # Extract all tables
            
            for table in tables:
                for row in table:
                    row_text = " ".join([cell.strip() for cell in row if cell])  # Combine row text
                    
                    # Debugging Step: Print each row
                    print(f"🔍 Checking Table Row: {row_text}")
                    
                    if "PO Number" in row_text:
                        match = re.search(r"PO Number\s*[:\-]?\s*(.+)", row_text, re.IGNORECASE)
                        if match:
                            po_number = match.group(1).strip()
                            print(f"✅ Found PO Number: {po_number}")
                            return po_number  # Return full PO Number
                        
    return "NOT FOUND"

def extract_invoice_details(text, pdf_path):
    """Extracts invoice details using regex and retrieves PO Number from tables."""

    # Debugging Step: Print extracted text
    print(f"\n📝 Extracted Text:\n{text}")

    invoice_number_match = re.search(r"Invoice Number:\s*(.+)", text, re.IGNORECASE)
    due_date_match = re.search(r"Due Date:\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    bill_to_match = re.search(r"Bill To:\s*(.+)", text, re.IGNORECASE)

    # Extract PO Number from the table inside the PDF
    po_number = extract_po_number_from_pdf(pdf_path)

    # If PO Number wasn't found in the table, try extracting from text
    if po_number == "NOT FOUND":
        po_number_match = re.search(r"Software Development services.*?([\w\-/\.]+)", text, re.IGNORECASE)
        po_number = po_number_match.group(1).strip() if po_number_match else "NOT FOUND"

    # Extract Total Amount Due
    total_amount_match = re.search(r"Total Amount Due\s*\$([\d,]+\.\d{2})", text, re.IGNORECASE)
    total_amount = total_amount_match.group(1).strip() if total_amount_match else "NOT FOUND"

    extracted_data = {
        "Invoice Number": invoice_number_match.group(1).strip() if invoice_number_match else "NOT FOUND",
        "Due Date": due_date_match.group(1).strip() if due_date_match else "NOT FOUND",
        "Bill To": bill_to_match.group(1).strip() if bill_to_match else "NOT FOUND",
        "PO Number": po_number,
        "Total Amount Due": total_amount
    }

    print(f"📊 Extracted Data: {extracted_data}")  # Debugging Step
    return extracted_data

# Process all PDFs
print(f"📂 Files in directory: {os.listdir(pdf_folder)}")
data_list = []

for filename in os.listdir(pdf_folder):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, filename)
        try:
            print(f"\nProcessing: {filename}")
            text = extract_text_from_pdf(pdf_path)
            extracted_data = extract_invoice_details(text, pdf_path)
            
            if any(value != "NOT FOUND" for value in extracted_data.values()):
                data_list.append(extracted_data)
            else:
                print(f"⚠️ No valid data found in {filename}")
        except Exception as e:
            print(f"❌ Error processing {filename}: {e}")

# Ensure output directory exists
os.makedirs(os.path.dirname(output_csv), exist_ok=True)

# Debugging Step: Check final extracted data
print(f"\n📝 Final Data List: {data_list}")

# Write to CSV only if valid data exists
if data_list:
    with open(output_csv, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["Invoice Number", "Due Date", "Bill To", "PO Number", "Total Amount Due"])
        writer.writeheader()
        writer.writerows(data_list)
    print(f"\n✅ Data successfully extracted and saved to {output_csv}")
else:
    print("\n⚠️ No valid data extracted. CSV file was not created.")