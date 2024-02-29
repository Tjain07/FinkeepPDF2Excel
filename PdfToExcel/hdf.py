import pdfplumber
import pandas as pd
import os

def process_pdf(pdf_file_path, output_folder):
    try:
        # Open the PDF file
        with pdfplumber.open(pdf_file_path) as pdf:
            all_data = []  # Initialize a list to store data line-wise

            # Extract text lines from each page starting from line 15
            for page in pdf.pages:
                text_lines = page.extract_text().split('\n')[14:]
                all_data.extend(text_lines)

        if all_data:
            # Create a pandas DataFrame with a single column named 'Data'
            df = pd.DataFrame(all_data, columns=['Data'])

            # Save the DataFrame to an Excel file in the specified output folder
            excel_file_name = os.path.splitext(os.path.basename(pdf_file_path))[0] + '_data_from_15th_line.xlsx'
            excel_file_path = os.path.join(output_folder, excel_file_name)
            df.to_excel(excel_file_path, index=False, header=False)  # Exclude header in each line

            print(f"Successfully processed: {pdf_file_path}")
        else:
            print(f"No data found in: {pdf_file_path}")

    except Exception as e:
        print(f"Error processing {pdf_file_path}: {e}")

# Define the folder path where your PDF files are located
folder_path = 'C:/Users/jaint/Downloads/PdfToExcel/PdfToExcel/Input/kotak'
output_folder = 'C:/Users/jaint/Downloads/PdfToExcel/PdfToExcel/Output/kotak'  # Specify your output folder path

# Get a list of all PDF files in the folder
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# Loop through each PDF file and process it
for pdf_file in pdf_files:
    pdf_file_path = os.path.join(folder_path, pdf_file)
    process_pdf(pdf_file_path, output_folder)
