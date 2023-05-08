import os

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


def get_pdf_pages_from_pub(pub_location=input_fileGlobal, save_pdf_location=pdf_folderGlobal):
    convert_exl_to_pdf(pub_location, save_pdf_location)
    


    
    

   





if __name__ == "__main__":
    convert_exl_to_pdf()
   