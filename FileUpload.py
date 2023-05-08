import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import ExporterPub as ConverterPDF
from PIL import Image, ImageTk
import tkinter.font as tkFont
from tkinter import ttk


class FileUploader(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Interface utilisateur")
        self.master.geometry('300x250')
        
        self.master.maxsize(300, 250)
        self.master.minsize(300, 250)
        self.master.config(bg='white')
        self.create_widgets()

    def create_widgets(self):

        style = ttk.Style()
        style.configure('TButton', font=('Segoe UI', 11), padding=5)
        style.configure('TLabel', font=('Segoe UI', 12), padding=10)
        style.configure('TFrame', background='white')

        # Frame pour le bouton de sélection de fichier
        self.button_frame = ttk.Frame(self.master)
        self.button_frame.pack(side="bottom", pady=10)

        self.upload_button = ttk.Button(self.button_frame, text="Escolha o Ficheiro", command=self.upload_file)
        self.upload_button.pack(side="left", padx=10)


        self.file_image = ttk.Label(self.master, anchor='center',style="" ,background=None)
        self.file_image.pack(pady=30)

        self.confirm_button = ttk.Button(self.button_frame, text="Confirmar", command=self.runcode)
        self.confirm_button.pack(side="left", padx=10)



    def upload_file(self):
        global file_path
        file_path = filedialog.askopenfilename()
        print("Le fichier sélectionné est :", file_path)
        _, file_extension = os.path.splitext(file_path)
        if file_extension == ".pub":
            image_path = os.path.abspath("PubExporterPdf/assets/Pub.png")
            """ image_path = "C:\\Users\\josel\\OneDrive\\Documents\\GitHub\\Hivy\\PubExporterPdf\\assets\\Pub.png"  """
            image = Image.open(image_path)
            image = image.resize((100, 100), Image.ANTIALIAS) 
            photo = ImageTk.PhotoImage(image)
            self.file_image.config(image=photo)
            #self.text_Path.config(text=file_path)
            self.file_image.image = photo 
            #self.text_Path.text = file_path
        elif file_extension == ".pdf":
            image_path = os.path.abspath("PubExporterPdf/assets/Pdf.png")
            print(image_path)
            """ image_path = "assets/Pdf.png" """ 
            image = Image.open(image_path)
            image = image.resize((100, 100), Image.ANTIALIAS) 
            photo = ImageTk.PhotoImage(image)
            self.file_image.config(image=photo)
            #self.text_Path.config(text=file_path)
            self.file_image.image = photo 
            #self.text_Path.text = file_path
            

            

       

    def runcode(self):
        _, file_extension = os.path.splitext(file_path)
        if file_extension == ".pub":
            try:
                ConverterPDF.get_pdf_pages_from_pub(pub_location=file_path)
            except:
                messagebox.showerror('file_path esta :', 'Vazio')
        elif file_extension == ".pdf":
            try:
                ConverterPDF.cut_pdf_to_pages(path_PDF=file_path)
            except:
                messagebox.showerror('file_path esta :', 'Vazio')


