import pdfplumber
import pandas as pd
import os

def extract_column_names(pdf_path, page_number=0, table_index=1):
    column_names = []

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number]
        table = page.extract_tables()[table_index]

        if table:
            # Assuming the first row contains column names
            column_names = table[0]

    return column_names

def process_pdf(pdf_file_path, output_folder):
    try:
        # Open the PDF file
        with pdfplumber.open(pdf_file_path) as pdf:
            all_tables = []  # Initialize a list to store tables from all pages

            for page in pdf.pages:
                # Extract tables from the current page
                page_tables = page.extract_tables()

                if page_tables:
                    all_tables.extend(page_tables)

        # Combine all the tables from different pages into a single list
        all_data = []
        for table in all_tables:
            column_names = extract_column_names(pdf_file_path, page_number=0, table_index=1)
            for row in table[0:]:
                # Ensure that the row has the same number of columns as the column names
                if len(row) == len(column_names):
                    all_data.append(row)

        if all_data:
            # Create a pandas DataFrame with dynamically determined column names
            df = pd.DataFrame(all_data, columns=column_names)

            # Save the DataFrame to an Excel file in the specified output folder
            excel_file_name = os.path.splitext(os.path.basename(pdf_file_path))[0] + '.xlsx'
            excel_file_path = os.path.join(output_folder, excel_file_name)
            df.to_excel(excel_file_path, index=False)

            print(f"Successfully processed: {pdf_file_path}")
        else:
            print(f"No data found in: {pdf_file_path}")

    except Exception as e:
        print(f"Error processing {pdf_file_path}: {e}")

# Define the folder path where your PDF files are located
folder_path = 'PdfToExcel/Input/IDFC'
output_folder = 'PdfToExcel/Output/IDFC'  # Specify your output folder path

# Get a list of all PDF files in the folder
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# Loop through each PDF file and process it
for pdf_file in pdf_files:
    pdf_file_path = os.path.join(folder_path, pdf_file)
    process_pdf(pdf_file_path, output_folder)
