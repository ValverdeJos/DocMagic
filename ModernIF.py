import os
import tkinter
from tkinter import filedialog

import customtkinter
from PIL import Image,ImageTk
from tkinter import messagebox

import ExporterPub as ConverterPDF
import ExporterExecl as ConverterExcel


customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
extention=[".xls",".xlsm",".xlsx",".xlt",".xlsb",".xltx",".xltm"]


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("DocMagic")
        self.geometry(f"{1100}x{580}")
        self.image_path = ""
        self.icon_image = tkinter.PhotoImage(file=self.image_path)
        self.checkbox_var = customtkinter.BooleanVar()
        self.checkbox_var.set(True)

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)
        

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="DocMagic",font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame ,command=self.sidebar_button_event_Pub ,text="Publisher")
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_event_Excel, text="Excel")
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        self.sidebar_button_3 = customtkinter.CTkButton(self.sidebar_frame, command=self.sidebar_button_event_Pdf, text="Pdf")
        self.sidebar_button_3.grid(row=3, column=0, padx=20, pady=10)


        #DarkMode & LightMode
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))

        #Scaling
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # create checkbox and switch frame
        self.checkbox_1 = customtkinter.CTkCheckBox(master=self.sidebar_frame,text="Pdf Separados ON",variable=self.checkbox_var,onvalue=True, offvalue=False,command=self.CheckText)
        self.checkbox_1.grid(row=4, column=0, pady=(20, 0), padx=20, sticky="n")


        # set default values
        self.checkbox_1.select()
        self.appearance_mode_optionemenu.set("Dark")
        self.scaling_optionemenu.set("100%")


        self.Main_frame = customtkinter.CTkFrame(self)
        self.Main_frame.grid(row=0, column=1, rowspan=3 ,padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.logo = customtkinter.CTkLabel(self.Main_frame, image=self.icon_image, text="")
        self.logo.place(relx=0.5, rely=0.2, anchor="n")



        self.Main_button_1 = customtkinter.CTkButton(self.Main_frame ,text="Upload",width=120, height=40, font=customtkinter.CTkFont(size=16, weight="bold"),command=self.UploadFiles)
        self.Main_button_1.place(relx=0.35, rely=0.9,anchor="s")
        self.Main_button_2 = customtkinter.CTkButton(self.Main_frame ,text="Confirme",width=120, height=40, font=customtkinter.CTkFont(size=16, weight="bold"),command=self.RunCode)
        self.Main_button_2.place(relx=0.65, rely=0.9,anchor="s")
        


    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def sidebar_button_event_Pub(self):
        self.image_path="C:/Users/josel/Desktop/PubExporterPdf/assets/Pub.png"
        image = Image.open(self.image_path)
        image = image.resize((200, 200), Image.ANTIALIAS) 
        photo = ImageTk.PhotoImage(image)
        self.logo.configure(image=photo)
        self.logo.image = photo

    def sidebar_button_event_Excel(self):
        self.image_path="C:/Users/josel/Desktop/PubExporterPdf/assets/Xls.png"
        image = Image.open(self.image_path)
        image = image.resize((200, 200), Image.ANTIALIAS) 
        photo = ImageTk.PhotoImage(image)
        self.logo.configure(image=photo)
        self.logo.image = photo

    def sidebar_button_event_Pdf(self):
        self.image_path="C:/Users/josel/Desktop/PubExporterPdf/assets/Pdf.png"
        image = Image.open(self.image_path)
        image = image.resize((200, 200), Image.ANTIALIAS) 
        photo = ImageTk.PhotoImage(image)
        self.logo.configure(image=photo)
        self.logo.image = photo
        
    def UploadFiles(self):
        global file_path
        file_path = filedialog.askopenfilename()

    def RunCode(self):
        _, file_extension = os.path.splitext(file_path)
        
        if file_extension == ".pub":
            if self.checkbox_var.get() == True:

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

        elif file_extension in extention :

            if self.checkbox_var.get() == True:

                if file_path:
                    try:
                        ConverterExcel.get_pdf_pages_from_Excel(pub_location=file_path)
                    except:
                        messagebox.showerror('file_path esta :', 'Vazio')
                else:
                    messagebox.showerror('Erro', 'Não é um fichiero Excel.')

            else:
                if file_path:
                    try:
                        ConverterExcel.get_pdf_SimplePage_from_Excel()
                    except:
                        messagebox.showerror('file_path esta :', 'Vazio')
                else:
                    messagebox.showerror('Erro', 'Não é um fichiero Excel.')

    def CheckText(self):
        print(self.checkbox_var.get())
        if self.checkbox_var.get() == False:
            self.checkbox_1.configure(text='Pdf Separados OFF')
        else:
            self.checkbox_1.configure(text='Pdf Separados ON')
        

