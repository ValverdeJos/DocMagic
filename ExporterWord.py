import os
from PyPDF2 import PdfReader, PdfWriter
from win32com import client


desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
os.chdir(desktop_path)

input_fileGlobal = desktop_path + ""
pdf_folderGlobal = desktop_path + "\\Pdfs_Exported"
pdf_Global = pdf_folderGlobal + "\\Global_Word_PDF.pdf"
output_folder = pdf_folderGlobal + "\\pages_Word_Pdf"

def convert_Word_to_pdf(input_file: str, pdf_folder: str) -> str:  
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    
    # Criar o nome do fichiero para o PDF
    pdf_name = "Global_Word_PDF.pdf"
    pdf_path = os.path.join(pdf_folder, pdf_name)

    # Opening Microsoft Word
    word = client.Dispatch("Word.Application")

    # Read Word File
    doc = word.Documents.Open(input_file)

    

    # Converting into PDF File
    export_path = os.path.join(pdf_folder, pdf_name)
    doc.SaveAs(export_path, FileFormat=17)

    doc.Close()
    word.Quit()

    if os.path.exists(export_path):
        print("Le fichier PDF a été créé avec succès.")
    else:
        print("Une erreur s'est produite lors de la conversion en PDF.")
    
    return pdf_path

def cut_pdf_to_pages(PDFFile=pdf_Global):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    with open(PDFFile, 'rb') as pdf_file:
        # Criar un letor de Pdf
        pdf_reader = PdfReader(pdf_file)

        # Obter todas as Paginas
        num_pages = len(pdf_reader.pages)
        # print(num_pages)
        for page_num in range(num_pages):
            # Criar un escritor de Pdf
            pdf_writer = PdfWriter()

            # Obter a pagina
            page = pdf_reader.pages[page_num]

            
            # Adicionar a pagina o escritor PDF
            pdf_writer.add_page(page)

            # Criar um Novo Ficheiro PDF
            new_file_name = 'page' + str(page_num) + '.pdf'
            new_file = open(os.path.join(output_folder, new_file_name), 'wb')
            Way_single_pdf = f"{output_folder}\\{new_file_name}"
            pdf_writer.write(new_file)
            new_file.close()

    
            
            # print(page_num)

def get_pdf_pages_from_Word(pub_location = input_fileGlobal, save_pdf_location = pdf_folderGlobal):
    path_pdf = convert_Word_to_pdf(input_file = pub_location,pdf_folder = save_pdf_location)
    cut_pdf_to_pages(PDFFile=path_pdf)

def get_pdf_SimplePage_from_Word(pub_location=input_fileGlobal, save_pdf_location=pdf_folderGlobal):
    convert_Word_to_pdf(pub_location, save_pdf_location)
  

