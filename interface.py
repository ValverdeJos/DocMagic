import tkinter as tk
from tkinter import filedialog
from tkinter import *
from FileUpload import FileUploader


app = tk.Tk()
app.title("Interface utilisateur")
app.geometry('300x300')
app.maxsize(300, 300)
app.minsize(300, 300)
app = FileUploader(master=app)

app.mainloop()






