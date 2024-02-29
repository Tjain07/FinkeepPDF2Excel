import pdfplumber
import openpyxl

def pdf_to_excel(pdf_path, excel_path):
    """Converts tables from a PDF file to an Excel file, stopping at the last page with a table."""

    with pdfplumber.open(pdf_path) as pdf:
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        last_page_with_table = None  # Keep track of the last page with a table

        for page in pdf.pages:
            try:
                table = page.extract_table()

                if table:  # Table found
                    last_page_with_table = page.page_number  # Update last page with table

                    for row in table:
                        sheet.append(row)

            except Exception as e:
                print(f"Error extracting table from page {page.page_number}: {e}")

        if last_page_with_table is None:
            print("No table found in the PDF.")
        else:
            workbook.save(excel_path)

# Example usage:
pdf_path = "PdfToExcel/Input/BOM/BOM_Statement_FTP_00005_xxxxxxxx2238_20210401_20220331_20221115022539.PDF" # Replace with your PDF file path
excel_path = "PdfToExcel/Output/BOM/BOM.xlsx" # Replace with your desired Excel file path
pdf_to_excel(pdf_path, excel_path)

