import pdfplumber

def extract_column_names(pdf_path, page_number=0, table_index=0):
    column_names = []

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number]
        table = page.extract_tables()[table_index]

        if table:
            # Assuming the first row contains column names
            column_names = table[0]

    return column_names

# Example usage:
pdf_path = 'C:\\Users\\jaint\\Downloads\\PdfToExcel\\PdfToExcel\\Input\\IDBI\\CFI-IDBI Bank Ac.pdf'
column_names = extract_column_names(pdf_path)

print("Column Names:", column_names)
