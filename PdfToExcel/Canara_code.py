import pdfplumber
import pandas as pd
import os
import re

# Define the folder path where your PDF files are located
folder_path = 'PdfToExcel/Input/Canara'

# Get a list of all PDF files in the folder
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# Function to extract tables and match the column names
def extract_and_rename_tables(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []  # Initialize a list to store tables from all pages
        # Assuming that the column names are in the first row of the first page
        first_page = pdf.pages[0]
        first_row_text = first_page.extract_text().split('\n')[16]
        
        # You may need to further process the text based on your PDF's structure
        # For example, you might use regular expressions to extract potential column names
        
        # Split the text by spaces, commas, or other delimiters
        potential_column_names = re.split(r'\s|,|;|\t', first_row_text)
        
        # Combine the first four columns into two columns
        combined_column_names = [potential_column_names[0] + ' ' + potential_column_names[1],
                                 potential_column_names[2] + ' ' + potential_column_names[3]]
        
        # Continue combining adjacent duplicate column names
        for name in potential_column_names[4:]:
            if name != combined_column_names[-1]:
                combined_column_names.append(name)
        
        # Create a DataFrame with a single row using the combined column names
        column_df = pd.DataFrame([combined_column_names], columns=combined_column_names)

        if not column_df.empty:
            column_names_array = column_df.columns.to_numpy()

        for page_number, page in enumerate(pdf.pages):
            # Skip the first 8 rows of the first page (page_number == 0)
            if page_number == 0:
                rows_to_skip = 1
            else:
                rows_to_skip = 0

            # Extract tables from the current page, skipping the specified rows
            page_tables = page.extract_tables()

            if page_tables:
                # Skip rows if necessary
                page_tables = page_tables[rows_to_skip:]

                all_tables.extend(page_tables)

        # Combine all the tables from different pages into a single list
        all_data = []
        for table in all_tables:
            # Use the first row as column headers
            data_rows = table[1:]
            all_data.extend(data_rows)

        return pd.DataFrame(all_data, columns=column_names_array)

# Loop through each PDF file and process it
for pdf_file in pdf_files:
    pdf_file_path = os.path.join(folder_path, pdf_file)

    try:
        df = extract_and_rename_tables(pdf_file_path)

        # Save the DataFrame to an Excel file with a name based on the PDF file
        excel_file_name = pdf_file.replace('.pdf', '.xlsx')
        excel_file_path = os.path.join(folder_path, excel_file_name)
        df.to_excel(excel_file_path, index=False)

        print(f"Successfully processed: {pdf_file}")
    except Exception as e:
        # Print the error message without raising it
        print(f"Error processing {pdf_file}: {e}")
