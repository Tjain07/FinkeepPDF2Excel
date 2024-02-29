'''import pdfplumber
import pandas as pd

# Open the PDF file
with pdfplumber.open('icici.pdf') as pdf:
    all_tables = []  # Initialize a list to store tables from all pages

    for page in pdf.pages:
        # Extract tables from the current page #check other methode to etract table
        page_tables = page.extract_tables()

        if page_tables:
            all_tables.extend(page_tables)

# Combine all the tables from different pages into a single list
all_data = []
for table in all_tables:
    all_data.extend(table)

# Create a pandas DataFrame
df = pd.DataFrame(all_data, columns=["Sr No", "Value Date", "Transaction Date", "ChequeNumber", "Transaction Remarks", "DebitAmount", "CreditAmount", "Balance(INR)" ])

# Save to Excel
df.to_excel("HDFC.xlsx", index=False)


#df = pd.DataFrame(all_data, columns=["Date", "Narration", "Chq./Ref.No.", "Value Dt", "Withdrawal Amt.", "Deposit Amt.", "Closing Balance" ])'''

import pdfplumber
import pandas as pd

# Open the PDF file
with pdfplumber.open('KOTAK BANK.pdf') as pdf:
    all_tables = []  # Initialize a list to store tables from all pages

    for page in pdf.pages:
        # Extract tables from the current page
        page_tables = page.extract_tables()

        if page_tables:
            all_tables.extend(page_tables)

# Define a function to determine column names
def determine_column_names(data):
    if data:
        # Assume the first row contains the column names
        column_names = data[0]
    else:
        column_names = None
    return column_names

# Check if there are tables in all_tables
if all_tables:
    # Determine column names from the first table
    column_names = determine_column_names(all_tables[0])
    # Create a pandas DataFrame with inferred column names
    df = pd.DataFrame(all_tables[1:], columns=column_names)
    # Save to Excel
    df.to_excel("HDFC.xlsx", index=False)
else:
    print("No tables found in the PDF.")
