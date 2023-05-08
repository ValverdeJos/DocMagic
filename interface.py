import tkinter as tk
from tkinter import filedialog
from tkinter import *
from FileUpload import FileUploader


app = tk.Tk()

app = FileUploader(master=app) 
app.mainloop()






