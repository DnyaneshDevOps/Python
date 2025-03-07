import pdfplumber
import pandas as pd
import re

def extract_invoice_data(pdf_path):
    extracted_data = []
    
    pdf_path = r"D:\Code\Invoices\invoice002-vai-gitex.pdf"
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                invoice_number = re.search(r'Invoice Number[:\s]+(\S+)', text)
                invoice_date = re.search(r'Due Date[:\s]+(\S+)', text)
                billed_to = re.search(r'Bill To[:\s]+([\w\s]+)', text)
                po_number = re.search(r'Software Development services[:\s]+(\S+)', text)
                value = re.search(r'Total Amount Due:\s*\$?([\d,]+\.?\d*)', text) ##re.search(r'Total Amount Due[:\s]+([\d,.]+)', text)
                
                data = {
                    "Invoice Number": invoice_number.group(1) if invoice_number else "N/A",
                    "Invoice Date": invoice_date.group(1) if invoice_date else "N/A",
                    "Billed To": billed_to.group(1).strip() if billed_to else "N/A",
                    "PO Number": po_number.group(1) if po_number else "N/A",
                    "Value": value.group(1) if value else "N/A"
                }
                extracted_data.append(data)
    
    return extracted_data

def save_to_csv(data, output_csv):
    df = pd.DataFrame(data)
    df.to_csv(output_csv, index=False)

if __name__ == "__main__":
    pdf_path = "input_invoice.pdf"  # Update with your PDF file path
    output_csv = "output.csv"
    data = extract_invoice_data(pdf_path)
    save_to_csv(data, output_csv)
    print(f"Extracted data saved to {output_csv}")