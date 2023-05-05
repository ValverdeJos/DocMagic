import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import ExporterPub as ConverterPDF

class FileUploader(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.upload_button = tk.Button(self)
        self.upload_button["text"] = "Choisir un fichier"
        self.upload_button["command"] = self.upload_file
        self.upload_button.pack(side="top")

    def upload_file(self):
        global file_path
        file_path = filedialog.askopenfilename()
        print("Le fichier sélectionné est :", file_path)
        

    def runcode():
        # votre code ici qui utilise le fichier choisi
        try:
            ConverterPDF.get_pdf_pages_from_pub(pub_location=file_path)
        except:
            messagebox.showerror('file_path esta :','Vazio')
        