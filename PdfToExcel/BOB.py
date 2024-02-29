import os
import pdfplumber
import pandas as pd
import easyocr
import numpy as np
from PIL import Image
import pdf2image

def extract_table_data(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        # Assuming the table is in the first page; change the page number accordingly
        page = pdf.pages[0]

        # Extract text using OCR from the entire page
        text = extract_text_with_ocr(page)

        # Split the text into lines and assume the first line contains column names
        lines = text.split('\n')
        column_names = lines[0].split()

        # Extract data from the remaining lines
        data = [line.split() for line in lines[1:]]

    return column_names, data

def extract_text_with_ocr(page):
    # Convert PDF page to an image using pdf2image
    image = page.to_image(resolution=300).original
    image_pil = Image.fromarray(np.array(image))

    # Convert PIL Image to NumPy array
    image_np = np.array(image_pil)

    # Use easyocr for OCR
    reader = easyocr.Reader(['en'])
    result = reader.readtext(image_np, detail=0)

    return ' '.join(result)

def convert_to_excel(column_names, data, excel_file):
    df = pd.DataFrame(data, columns=column_names)
    df.to_excel(excel_file, index=False)

def process_pdfs(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)
            output_file = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.xlsx")

            column_names, data = extract_table_data(pdf_path)
            convert_to_excel(column_names, data, output_file)

            print(f"Table extracted from {pdf_path} and saved to {output_file}")

if __name__ == "__main__":
    # Hardcoded input and output folders
    input_folder = "PdfToExcel/Input/Bank of baroda"
    output_folder = "PdfToExcel/Output/Bank of baroda"

    process_pdfs(input_folder, output_folder)
