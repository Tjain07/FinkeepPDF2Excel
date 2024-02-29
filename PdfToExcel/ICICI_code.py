import pdfplumber
import pandas as pd
import os

# Prompt the user to input the folder path where PDF files are located
folder_path = 'PdfToExcel/Input/IDBI'

# Prompt the user to input the output folder path where Excel files will be saved
output_folder = 'PdfToExcel/Output/IDBI'

# Get a list of all PDF files in the input folder
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# Function to determine column names based on the first row of the table
def determine_column_names(table):
    return table[0]

# Loop through each PDF file and process it
for pdf_file in pdf_files:
    pdf_file_path = os.path.join(folder_path, pdf_file)

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
            column_names = determine_column_names(table)
            for row in table[0:]:
                # Ensure that the row has the same number of columns as the column names
                if len(row) == len(column_names):
                    all_data.append(row)

        if all_data:
            # Create a pandas DataFrame with dynamically determined column names
            df = pd.DataFrame(all_data, columns=column_names)

            # Save the DataFrame to an Excel file in the specified output folder
            excel_file_name = pdf_file.replace('.pdf', '.xlsx')
            excel_file_path = os.path.join(output_folder, excel_file_name)
            df.to_excel(excel_file_path, index=False)
            
            print(f"Successfully processed: {pdf_file}")
        else:
            print(f"No data found in: {pdf_file}")

    except Exception as e:
        print(f"Error processing {pdf_file}: {e}")
