import os
import win32com.client
input_fileGlobal = "C:\\Users\\josel\\OneDrive\\Bureau\\ProjectExporterPub\\FICHASTECNICAS2023.pub"
pdf_folderGlobal = "C:\\Users\\josel\\OneDrive\\Bureau\\ProjectExporterPub\\PdfExported"

def convert_pub_to_pdf(input_file: str, pdf_folder: str) -> str:  
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    publisher = win32com.client.Dispatch("Publisher.Application")

    # Open the input file
    document = publisher.Open(input_file)

    # Export each page to a separate PDF file

    # Create a file name for the PDF
    pdf_name = "final.pdf"
    pdf_path = os.path.join(pdf_folder, pdf_name)

    # Export the current page to PDF
    try:
        document.ExportAsFixedFormat(2, pdf_path)
    except:
        pass

    # Close the document and Publisher application
    document.Close()
    publisher.Quit()
    
    return pdf_path

def cut_pdf_to_pages(pdf_path):
    pass

def get_pdf_pages_from_pub(pub_location=input_fileGlobal, save_pdf_location=pdf_folderGlobal):
    pdf_path = convert_pub_to_pdf(pub_location, save_pdf_location)
    cut_pdf_to_pages(pdf_path)


if __name__ == "__main__":
    get_pdf_pages_from_pub()