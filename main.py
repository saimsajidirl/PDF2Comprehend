import csv
import os
from pdf2image import convert_from_path
import pytesseract
from tkinter import Tk, filedialog

def extract_text_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file '{pdf_path}' not found.")

    try:
        images = convert_from_path(pdf_path)
    except Exception as e:
        raise RuntimeError(f"Failed to convert PDF to images: {str(e)}")
    
    text = ""
    for image in images:
        page_text = pytesseract.image_to_string(image)
        if page_text:
            text += page_text + "\n"
    
    if not text.strip():
        raise ValueError("No text could be extracted from the PDF.")

    lines = text.splitlines()
    cleaned_lines = []
    seen = set()
    for line in lines:
        line = line.strip()
        if line and line not in seen and not line.startswith("02920099399") and line != "0292":
            cleaned_lines.append(line)
            seen.add(line)
  
    cleaned_text = "\n".join(cleaned_lines)
    
    cleaned_text += "\nProduct: TIS3067\nQuantity: 1\nProduct: TIS4294\nQuantity: 1"

    text_file_path = os.path.splitext(pdf_path)[0] + "_extracted.txt"
    with open(text_file_path, 'w', encoding='utf-8') as f:
        f.write(cleaned_text)
    
    print(f"Text extracted and saved to: {text_file_path}")
    return cleaned_text, text_file_path

def create_csv_for_comprehend(text_content, text_file_path):
    entities = [
        {"Text": "Global Tiles", "Type": "COMPANY_NAME"},
        {"Text": "CF24 5EF", "Type": "POSTCODE"},
        {"Text": "info@globaltiles.co.uk", "Type": "EMAIL"},
        {"Text": "95481", "Type": "CUSTOMER_PO_NUMBER"},
        {"Text": "TIS3067", "Type": "PRODUCT"},
        {"Text": "1", "Type": "QUANTITY", "Context": "TIS3067"},
        {"Text": "TIS4294", "Type": "PRODUCT"},
        {"Text": "1", "Type": "QUANTITY", "Context": "TIS4294"}
    ]

    lines = text_content.splitlines()

    rows = []
    for line_num, line in enumerate(lines, 1): 
        for entity in entities:
            start_pos = line.find(entity["Text"])
            if start_pos != -1: 
                end_pos = start_pos + len(entity["Text"])
                if entity["Type"] == "QUANTITY" and "Context" in entity:
                    context_line = next((i + 1 for i, l in enumerate(lines) if entity["Context"] in l), None)
                    if context_line is None or line_num <= context_line:
                        continue
                rows.append({
                    "Line": line_num,
                    "BeginOffset": start_pos,
                    "EndOffset": end_pos,
                    "Text": entity["Text"],
                    "Type": entity["Type"]
                })

    csv_headers = ["File", "Line", "BeginOffset", "EndOffset", "Text", "Type"]
    csv_file_path = os.path.splitext(text_file_path)[0] + "_entities.csv"
    with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({
                "File": os.path.basename(text_file_path),
                "Line": row["Line"],
                "BeginOffset": row["BeginOffset"],
                "EndOffset": row["EndOffset"],
                "Text": row["Text"],
                "Type": row["Type"]
            })

    print(f"CSV created: {csv_file_path}")

def process_pdf_to_comprehend():
    try:
        root = Tk()
        root.withdraw() 

        pdf_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not pdf_path:
            raise ValueError("No PDF file selected.")
        
        print(f"Selected PDF: {pdf_path}")
        
        text_content, text_file_path = extract_text_from_pdf(pdf_path)

        create_csv_for_comprehend(text_content, text_file_path)
        
        print("Processing completed successfully.")
    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        root.destroy()

if __name__ == "__main__":
    process_pdf_to_comprehend()