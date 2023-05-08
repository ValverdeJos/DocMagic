import os
from PyPDF2 import  PdfReader, PdfWriter
import win32com.client 


import  jpype     
import  asposecells     
jpype.startJVM() 
from asposecells.api import Workbook

desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
os.chdir(desktop_path)

input_fileGlobal = "C:\\Users\\josel\\Desktop\\PubExporterPdf\\pinel.xls"
pdf_folderGlobal = desktop_path+"\\PdfsExportedToExel"
pdf_Global = pdf_folderGlobal+"\\GlobalPDFtoExel.pdf"
output_folder = pdf_folderGlobal +"\\pagesExel"


def convert_exl_to_pdf(input_file: str, pdf_folder: str) -> str:  
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    workbook = Workbook(input_file)
    workbook.save(pdf_Global)
    jpype.shutdownJVM()

def cut_pdf_to_pages(path_PDF=pdf_Global):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    with open(path_PDF, 'rb') as pdf_file:
        # Criar un letor de Pdf
        pdf_reader = PdfReader(pdf_file)

        # Obter todas as Paginas
        num_pages = len(pdf_reader.pages)
        print(num_pages)
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
            pdf_writer.write(new_file)
            new_file.close()

    
            print(page_num)
  

def get_pdf_pages_from_Excel(pub_location=input_fileGlobal, save_pdf_location=pdf_folderGlobal):
    convert_exl_to_pdf(pub_location, save_pdf_location)
    cut_pdf_to_pages()
    




   