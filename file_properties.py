#pip install xlrd
#pip install python-docx
#pip install pywin32

import xlrd
import openpyxl
from docx import Document
import win32com.client

def extract_excel_properties_xls(file_path):
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False  # Set to True if you want Excel to be visible

    try:
        workbook = excel_app.Workbooks.Open(file_path)
        
        # Access core properties
        core_properties = workbook.BuiltinDocumentProperties

        # print(dir(core_properties),'\n')

        # Print core properties
        print("\nXLS File Properties:\n---------------------------")
        print(f"Title: {core_properties('Title').Value}")
        print(f"Author: {core_properties('Author').Value}")
        print(f"Last Author: {core_properties('Last Author').Value}")
        # print(f"Created: {core_properties('Created').Value}")
        # print(f"Last Modified: {core_properties('Last Modified').Value}")

        print(type(core_properties))

    except Exception as e:
        print(f"Error: {e}")

    finally:
        # Close the workbook and quit Excel
        workbook.Close(False)
        excel_app.Quit()


def extract_excel_properties_xlsx(file_path):
    # Load the Excel workbook
    wb = openpyxl.load_workbook(file_path)

    # print(dir(wb.properties))
    # Get properties
    properties = {
        'Title': wb.properties.title,
        'Author': wb.properties.creator,
        'Created': wb.properties.created,
        'Last Modified By': wb.properties.lastModifiedBy,
        'Modified': wb.properties.modified,
    }

    # Print properties
    print("\nXLSX File Properties:\n---------------------------")
    for key, value in properties.items():
        print(f"{key}: {value}")

    # Close the workbook
    wb.close()


def extract_excel_properties_docx(file_path):
    # Open the Excel file as a Word document
    doc = Document(file_path)

    # Extract extended properties
    properties = {
        'Title': doc.core_properties.title,
        'Author': doc.core_properties.author,
        'Last Author': doc.core_properties.last_modified_by,
        'Created': doc.core_properties.created,
        'Modified': doc.core_properties.modified,
    }

    # Print properties
    print("\nDOCX File Properties:\n---------------------------")
    for key, value in properties.items():
        print(f"{key}: {value}")

    

if __name__ == "__main__":
    file_path_xls = "./doc_files/BB01B.xls"
    file_path_xls_full = r"C:\Users\janshimy\Downloads\BB01B.xls"
    file_path_xlsx = "I:\Projects\ABC Community\DataMgt\Reports for Team\CSPB Monthly Dashboard Reports\CSPB Dashboard 2023 10.xlsx"
    file_path_docx = "./doc_files/Haleon_Protocol_ Jul 19 2023_Final.docx"

    extract_excel_properties_xlsx(file_path_xlsx)

    extract_excel_properties_docx(file_path_docx)
    
    extract_excel_properties_xls(file_path_xls_full) # This function might need a full path
    
