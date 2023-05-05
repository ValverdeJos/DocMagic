import tkinter as tk
from tkinter import filedialog
from tkinter import *
from FileUpload import FileUploader


app = tk.Tk()
app.title("Interface utilisateur")
app.geometry('300x210')
app.maxsize(300, 210)
app.minsize(300, 210)
app = FileUploader(master=app)

app.mainloop()






