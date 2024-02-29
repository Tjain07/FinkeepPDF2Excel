import pdfplumber
import pandas as pd
import os

# Specify the input and output folder paths
input_folder = 'PdfToExcel\\Input\\SBI'
output_folder = 'PdfToExcel\\Output\\SBI'

# Get a list of all PDF files in the input folder
pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]

# Function to extract column names from the PDF
def extract_column_names(pdf_path, page_number=0):
    column_names = []

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number]
        text = page.extract_text()

        # Assuming that the column names are in the first row of the table
        first_row = text.split('\n')[19]
        all_columns = first_row.split()

        # Combine the first three columns
        combined_columns_1_3 = ' '.join(all_columns[:3])
        # Combine columns 5 and 6 together
        combined_columns_4 = ''.join(all_columns[3])
        # Combine columns 5 and 6 together
        combined_columns_5_6 = '_'.join(all_columns[4:6])
        # Keep the rest of the columns as they are
        remaining_columns = all_columns[6:]

        # Append the combined columns and the remaining columns to the list
        column_names.append(combined_columns_1_3)
        column_names.append(combined_columns_4)
        column_names.append(combined_columns_5_6)
        column_names.extend(remaining_columns)

    return column_names

# Function to extract tables and match the column names
def extract_and_rename_tables(pdf_path, column_names):
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []  # Initialize a list to store tables from all pages

        for page in pdf.pages:
            # Extract tables from the current page
            page_tables = page.extract_tables()

            if page_tables:
                all_tables.extend(page_tables)

        # Combine all the tables from different pages into a single list
        all_data = []
        for table in all_tables:
            if len(column_names) == len(table[0]):
                all_data.extend(table)

        return pd.DataFrame(all_data, columns=column_names)

# Loop through each PDF file in the input folder and process it
for pdf_file in pdf_files:
    pdf_file_path = os.path.join(input_folder, pdf_file)

    try:
        # Extract column names from the PDF
        column_names = extract_column_names(pdf_file_path, page_number=0)
        # Extract tables and match the column names
        df = extract_and_rename_tables(pdf_file_path, column_names)

        # Save the DataFrame to an Excel file in the specified output folder
        excel_file_name = pdf_file.replace('.pdf', '.xlsx')
        excel_file_path = os.path.join(output_folder, excel_file_name)
        df.to_excel(excel_file_path, index=False)

        print(f"Successfully processed: {pdf_file}")
    except Exception as e:
        # Print the error message without raising it
        print(f"Error processing {pdf_file}: {e}")
