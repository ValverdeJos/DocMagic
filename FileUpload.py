import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import ExporterPub as ConverterPDF
from PIL import Image, ImageTk
import tkinter.font as tkFont


class FileUploader(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.button_frame = tk.Frame(self)
        self.button_frame.pack(side="bottom", pady=10)

        self.upload_button = tk.Button(self.button_frame)
        self.upload_button["text"] = "Escolha o Ficheiro"
        self.upload_button["command"] = self.upload_file
        self.upload_button.pack(side="left", padx=10)

        self.file_image = tk.Label(self)
        self.file_image.pack(side="top", pady=20)

        """ self.text_Path = tk.Label(self)
        self.text_Path.pack() """

        self.confirm_button = tk.Button(self.button_frame)
        self.confirm_button["text"] = "Confirmer"
        self.confirm_button["command"] = self.runcode
        self.confirm_button.pack(side="left", padx=10)
        
        

        self.pack(side="bottom")

    def upload_file(self):
        global file_path
        file_path = filedialog.askopenfilename()
        print("Le fichier sélectionné est :", file_path)
        _, file_extension = os.path.splitext(file_path)
        if file_extension == ".pub":
            image_path = "C:\\Users\\josel\\OneDrive\\Documents\\GitHub\\Hivy\\PubExporterPdf\\assets\\Pub.png" # remplacez "chemin/vers/icon/pub.png" par le chemin vers l'icône du fichier pub
            image = Image.open(image_path)
            image = image.resize((100, 100), Image.ANTIALIAS) # Redimensionner l'image
            photo = ImageTk.PhotoImage(image)
            self.file_image.config(image=photo)
            #self.text_Path.config(text=file_path)
            self.file_image.image = photo # Garder une référence à l'image pour éviter qu'elle ne soit effacée par le garbage collector
            #self.text_Path.text = file_path
        elif file_extension == ".pdf":
            image_path = "C:\\Users\\josel\\OneDrive\\Documents\\GitHub\\Hivy\\PubExporterPdf\\assets\\Pdf.png" # remplacez "chemin/vers/icon/pub.png" par le chemin vers l'icône du fichier pub
            image = Image.open(image_path)
            image = image.resize((100, 100), Image.ANTIALIAS) # Redimensionner l'image
            photo = ImageTk.PhotoImage(image)
            self.file_image.config(image=photo)
            #self.text_Path.config(text=file_path)
            self.file_image.image = photo # Garder une référence à l'image pour éviter qu'elle ne soit effacée par le garbage collector
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


