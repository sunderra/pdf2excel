import streamlit as st
import os
from spire.xls import *
from spire.xls.common import *
from spire.pdf.common import *
from spire.pdf import *

def file_selector(folder_path='.'):
    filenames = os.listdir(folder_path)
    selected_filename = st.selectbox('Select a file', filenames)
    return os.path.join(folder_path, selected_filename)


# --- Excel to PDF ---
def excel_to_pdf(excel_path, pdf_path):
    try:
        # Create a Workbook object
        workbook = Workbook()

        # Load an Excel document from file
        workbook.LoadFromFile(excel_path)

        # Iterate through the worksheets in the workbook
        for sheet in workbook.Worksheets:

            # Get the PageSetup object
            pageSetup = sheet.PageSetup

            # Set page margins
            pageSetup.TopMargin = 0.3
            pageSetup.BottomMargin = 0.3
            pageSetup.LeftMargin = 0.3
            pageSetup.RightMargin = 0.3

        # Set worksheet to fit to page when converting
        workbook.ConverterSetting.SheetFitToPage = True

        # Convert all worksheets to PDF file
        workbook.SaveToFile(pdf_path, FileFormat.PDF)

        # Close the workbook
        workbook.Dispose()
    except Exception as e:
        print(f"Error during Excel to PDF conversion: {e}")


# --- PDF to Excel ---
def pdf_to_excel(pdf_path, excel_path):
    try:
        # Create a PdfDocument object
        pdf = PdfDocument()

        # Load a PDF document
        pdf.LoadFromFile(pdf_path)

        # Create an XlsxLineLayoutOptions object to specify the conversion options
        # Parameters: convertToMultipleSheet, rotatedText, splitCell, wrapText, overlapText
        convertOptions = XlsxLineLayoutOptions(True, True, False, True, False)

        # Set the conversion options
        pdf.ConvertOptions.SetPdfToXlsxOptions(convertOptions)

        # Save the PDF document to Excel XLSX format
        pdf.SaveToFile(excel_path, FileFormat.XLSX)

        # Close the PDF document
        pdf.Close()
    except Exception as e:
        print(f"Error during PDF to Excel conversion: {e}")

# --- Main execution ---

if __name__ == "__main__":
    filenames = file_selector("c:\.")

    uploaded_files = st.sidebar.file_uploader(
        "Choose one or more Excel files", type=['.xls','.xlsx'], accept_multiple_files=True
        )

    for uploaded_file in uploaded_files:
        st.sidebar.write("You selected filenames:", uploaded_file.name)
        # Convert Excel to PDF
        pdf_filename = os.path(uploaded_file.name).split('.')[0] + '.pdf'
        st.sidebar.write("Converting to :", pdf_filename)
        #excel_to_pdf(uploaded_file.name, pdf_filename)

    # Convert PDF to Excel
    #pdf_to_excel(pdf_file, excel_file)


