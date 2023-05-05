import tkinter as tk
from tkinter import filedialog
from tkinter import *
from FileUpload import FileUploader


app = tk.Tk()
app.title("Interface utilisateur")
app.geometry('300x300')
app = FileUploader(master=app)
confirm_button = Button(app, text="Confirmer", command=FileUploader.runcode)
confirm_button.pack(pady=10)
app.mainloop()