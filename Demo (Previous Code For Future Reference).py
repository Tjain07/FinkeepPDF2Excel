import os
import pdfplumber
import pandas as pd

# Set the input and output folder paths
input_folder = 'C:\\Users\\HP\\Desktop\\Various bank statements\\Input\\ICICI'
output_folder = 'C:\\Users\\HP\\Desktop\\Various bank statements\\Output'

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Iterate through all PDF files in the input folder
for pdf_filename in os.listdir(input_folder):
    if pdf_filename.endswith('.pdf'):
        pdf_path = os.path.join(input_folder, pdf_filename)
        
        # Open the PDF file
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
            all_data.extend(table)

        # Use the first row of the table as column names
        if all_data:
            columns = all_data[0]
            data = all_data[1:]

            # Create a pandas DataFrame with the extracted data and detected column names
            df = pd.DataFrame(data, columns=columns)

            # Define the output Excel filename based on the PDF filename
            output_filename = os.path.splitext(pdf_filename)[0] + '.xlsx'
            output_path = os.path.join(output_folder, output_filename)

            # Save to Excel
            df.to_excel(output_path, index=False)

print("All PDFs processed, and Excel files saved in the output folder.")
