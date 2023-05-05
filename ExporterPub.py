import os
import win32com.client
from PyPDF2 import  PdfReader, PdfWriter

input_fileGlobal = "C:\\Users\\josel\\OneDrive\\Bureau\\ProjectExporterPub\\FICHASTECNICAS2023.pub"
output_folder = "C:\\Users\\josel\\OneDrive\\Bureau\\PdfExported\\pages"
pdf_folderGlobal = "C:\\Users\\josel\\OneDrive\\Bureau\\PdfExported"
pdf_Global = "C:\\Users\\josel\\OneDrive\\Bureau\\PdfExported\\final.pdf"

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

def cut_pdf_to_pages(path_PDF):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    with open(path_PDF, 'rb') as pdf_file:
        # Créer un lecteur PDF
        pdf_reader = PdfReader(pdf_file)

        # Obtenir le nombre de pages
        num_pages = len(pdf_reader.pages)
        print(num_pages)
        for page_num in range(num_pages):
            # Créer un écrivain PDF
            pdf_writer = PdfWriter()

            # Obtenir la page
            page = pdf_reader.pages[page_num]

            # Ajouter la page à l'écrivain PDF
            pdf_writer.add_page(page)

            # Créer le nouveau fichier PDF
            new_file_name = 'page' + str(page_num) + '.pdf'
            new_file = open(os.path.join(output_folder, new_file_name), 'wb')
            Way_single_pdf = f"{output_folder}\\{new_file_name}"
            pdf_writer.write(new_file)
            new_file.close()

    
            extra_Text_Pdf(Way_single_pdf)
            print(page_num)
            
            

def extra_Text_Pdf(file):
    pdf_file_extra_Text_and_Image = open(file,'rb')
    pdf_reader_extra_Text_and_Image = PdfReader(pdf_file_extra_Text_and_Image)

    meta = pdf_reader_extra_Text_and_Image.metadata

    print(meta.title)
    text=''
    for page in range(1):
        text += pdf_reader_extra_Text_and_Image.pages[page].extract_text()
    print(text)
    
    pdf_file_extra_Text_and_Image.close()
    


def get_pdf_pages_from_pub(pub_location=input_fileGlobal, save_pdf_location=pdf_folderGlobal):
    pdf_path = convert_pub_to_pdf(pub_location, save_pdf_location)
    cut_pdf_to_pages(pdf_path)


if __name__ == "__main__":
    get_pdf_pages_from_pub()