import os
from pathlib import Path
import win32com.client
from PyPDF2 import  PdfReader, PdfWriter

desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
os.chdir(desktop_path)


input_fileGlobal = " "
pdf_folderGlobal = desktop_path+"\\PdfsExported"
pdf_Global = pdf_folderGlobal+"\\GlobalPDF.pdf"
output_folder = pdf_folderGlobal +"\\pages"

def convert_pub_to_pdf(input_file: str, pdf_folder: str) -> str:  
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    publisher = win32com.client.Dispatch("Publisher.Application")

    # Abrir o Ficheiro de entrada
    document = publisher.Open(input_file)

    
    # Criar o nome do fichiero para o PDF
    pdf_name = "GlobalPDF.pdf"
    pdf_path = os.path.join(pdf_folder, pdf_name)

    # exportar a pagina currente do PDF
    try:
        document.ExportAsFixedFormat(2, pdf_path)
    except:
        print("Error in Exporting Pdf")

   
    # Fechar o Document e o Publisher
    document.Close()
    publisher.Quit()
    
    return pdf_path

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
            Way_single_pdf = f"{output_folder}\\{new_file_name}"
            pdf_writer.write(new_file)
            new_file.close()

    
            extra_Text_Pdf(Way_single_pdf)
            print(page_num)
            
def extra_Text_Pdf(file):
    text=''
    pdf_file_extra_Text_and_Image = open(file,'rb')
    pdf_reader_extra_Text_and_Image = PdfReader(pdf_file_extra_Text_and_Image)
    
    
    for page in range(1):
        text += pdf_reader_extra_Text_and_Image.pages[page].extract_text()
    first_line=text.splitlines()[0]
    name_Title_PDF= first_line.split("|")
    counter = 0
    while True:
        if counter == 0:
            Way_return_pdf = f"{output_folder}\\{name_Title_PDF[0]}.pdf"
        else:
            Way_return_pdf = f"{output_folder}\\{name_Title_PDF[0]}_{counter}.pdf"
        
        if not os.path.exists(Way_return_pdf):
            break
        counter += 1
    print(Way_return_pdf)
    
    pdf_file_extra_Text_and_Image.close()
    os.rename(file, Way_return_pdf)

def get_pdf_pages_from_pub(pub_location=input_fileGlobal, save_pdf_location=pdf_folderGlobal):
    pdf_path = convert_pub_to_pdf(pub_location, save_pdf_location)
    cut_pdf_to_pages(pdf_path)


if __name__ == "__main__":
    get_pdf_pages_from_pub()
    