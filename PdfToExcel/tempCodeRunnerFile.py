import pdfplumber

def extract_column_names(pdf_path, page_number=0):
    column_names = []

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number]
        text = page.extract_text()

        # Assuming that the column names are in the first row of the table
        first_row = text.split('\n')[19]
        all_columns = first_row.split()

        # Combine the first three columns
        combined_columns = ' '.join(all_columns[:3])
        # Keep the rest of the columns as they are
        remaining_columns = all_columns[3:]

        # Append the combined columns and the remaining columns to the list
        column_names.append(combined_columns)
        column_names.extend(remaining_columns)

    return column_names

# Replace 'your_pdf_file.pdf' with the actual path to your PDF file
pdf_path = 'C:\\Users\\jaint\\Downloads\\PdfToExcel\\PdfToExcel\\Input\\SBI\\sbi sitamarhi.pdf'

column_names = extract_column_names(pdf_path)

print("Column Names:", column_names)
