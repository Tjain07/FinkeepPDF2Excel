import os
import pdfplumber
import pandas as pd

# Set the input and output folder paths
input_folder = 'H:\FinKeep\PdfToExcel\HDFC_in'
output_folder = 'H:\FinKeep\PdfToExcel\HDFC_out'

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

        # Create a pandas DataFrame
        #df = pd.DataFrame(all_data, columns=["Sr No", "Value Date", "Transaction Date", "ChequeNumber", "Transaction Remarks", "DebitAmount", "CreditAmount", "Balance(INR)"])
        df = pd.DataFrame(all_data, columns=["Date", "Narration", "Chq./Ref.No.", "Value Dt", "Withdrawal Amt.", "Deposit Amt.", "Closing Balance" ])

        # Define the output Excel filename based on the PDF filename
        output_filename = os.path.splitext(pdf_filename)[0] + '.xlsx'
        output_path = os.path.join(output_folder, output_filename)

        # Save to Excel
        df.to_excel(output_path, index=False)

print("All PDFs processed, and Excel files saved in the output folder.")
