import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox





import ExporterPub as ConverterPDF
#import ExporterExecl as ConverterExcel
from PIL import Image, ImageTk
#import tkinter.font as tkFont
from tkinter import ttk


#extention=[".xls",".xlsm",".xlsx",".xlt",".xlsb",".xltx",".xltm"]

class FileUploader(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Interface utilisateur")
        self.master.geometry('300x250')
        self.checkbox_var = tk.BooleanVar()
        self.checkbox_var.set(False)
        
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

        self.button_frame_frame = ttk.Frame(self.button_frame)
        self.button_frame_frame.pack(side="top", pady=10)   

        self.checkbox = ttk.Checkbutton(self.button_frame_frame,text='Separado por paginas',command=self.check_changed, variable=self.checkbox_var,onvalue=True, offvalue=False)
        self.checkbox.pack(padx=10)

        self.upload_button = ttk.Button(self.button_frame, text="Escolha o Ficheiro", command=self.upload_file)
        self.upload_button.pack(side="left", padx=10)

        
        self.file_image = ttk.Label(self.master, anchor='center',style="" ,background=None)
        self.file_image.pack(side="top",pady=10)

        self.confirm_button = ttk.Button(self.button_frame, text="Confirmar", command=self.runcode)
        self.confirm_button.pack(side="left", padx=10)
 

    def upload_file(self):
        global file_path
        file_path = filedialog.askopenfilename()
        # print("Le fichier sélectionné est :", file_path)
        _, file_extension = os.path.splitext(file_path)
        if file_extension == ".pub":
            image_path = os.path.abspath("PubExporterPdf/assets/Pub.png")
            image = Image.open(image_path)
            image = image.resize((100, 100), Image.ANTIALIAS) 
            photo = ImageTk.PhotoImage(image)
            self.file_image.config(image=photo)
            self.file_image.image = photo 
            
        elif file_extension == ".pdf":
            image_path = os.path.abspath("PubExporterPdf/assets/Pdf.png")
            # print(image_path)
            
            image = Image.open(image_path)
            image = image.resize((100, 100), Image.ANTIALIAS) 
            photo = ImageTk.PhotoImage(image)
            self.file_image.config(image=photo)
            self.file_image.image = photo 
            
        """  elif file_extension in extention:
            image_path = os.path.abspath("PubExporterPdf/assets/XLs.png")
            # print(image_path)
            #image_path = "assets/Pdf.png" 
            image = Image.open(image_path)
            image = image.resize((100, 100), Image.ANTIALIAS) 
            photo = ImageTk.PhotoImage(image)
            self.file_image.config(image=photo)
            #self.text_Path.config(text=file_path)
            self.file_image.image = photo 
            #self.text_Path.text = file_path """
            

            

    def check_changed(self):
        pass
        

    def runcode(self):
        _, file_extension = os.path.splitext(file_path)
        
        if file_extension == ".pub":
            if self.checkbox_var.get() == 1:

                try:
                    ConverterPDF.get_pdf_pages_from_pub(pub_location=file_path)
                    messagebox.showinfo("Sucesso","Converção em Pdfs Separados foi um sucesso")
                except:
                    messagebox.showerror('file_path esta :', 'Vazio')
            else:

                try:
                    ConverterPDF.get_pdf_SimplePage_from_pub(pub_location=file_path)
                    messagebox.showinfo("Sucesso","Converção em Pdf foi um sucesso")
                except:
                    messagebox.showerror('file_path esta :', 'Vazio')

        elif file_extension == ".pdf":

            try:
                ConverterPDF.cut_pdf_to_pages(path_PDF=file_path)
            except:
                messagebox.showerror('file_path esta :', 'Vazio')

        """         elif file_extension in extention :
            if file_path:
                try:
                    ConverterExcel.get_pdf_pages_from_Excel(pub_location=file_path)
                except:
                    messagebox.showerror('file_path esta :', 'Vazio')
            else:
                messagebox.showerror('Erro', 'Não é um fichiero Excel.') """




