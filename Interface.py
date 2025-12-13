import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from From_Spreadsheet import From_Spreadsheet
from ttkthemes import ThemedTk

def change_size():
   app.update()
   instr_lbl.config(wraplength=(app.winfo_reqwidth())-10)


# Create window
app = ThemedTk(theme='radiance')
app.title("SHL Metadata Creator - Alma Digital")
app.minsize(630, 220)
app.geometry('850x450')
#ttk.Style().theme_use('Breeze')
app.option_add("*Label*Background", "#f5f3f1")
app.option_add("*Radiobutton*Background", "#f5f3f1")
app.option_add('**Background', "#f5f3f1")
app.option_add('*Entry*Background', '#FFFFFF')
app.option_add('*ComboBox*Background', '#FFFFFF')
app['bg'] = "#f5f3f1"
app.configure(bg="#f5f3f1")

app.grid_columnconfigure(0, weight=1)
app.grid_columnconfigure(1, weight=0)
app.grid_rowconfigure(0, weight=0)
app.grid_rowconfigure(1, weight=1)

title_lbl = tk.Label(master=app, text="SHL Metadata Creator - Alma Digital", anchor="w", font = (20))
title_lbl.grid(row=0, column = 0, sticky="nw")

app_footer = tk.Label(master=app, text='Created by Arianna Milighetti, Digitisation Coordinator, Senate House Library. Documentation available on Digitisation SharePoint site. Created June 2025.', anchor='w', font=('Arial', 8))
app_footer.grid(row=2, column=0, sticky='nwse')

frame = ttk.Frame(app)
frame.grid(row=1, column=0, sticky='nwse')
frame.columnconfigure(0, weight=1)
frame.rowconfigure(0, weight=1)

canvas = tk.Canvas(frame)
scrollbarv = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
scrollbarh = ttk.Scrollbar(frame, orient="horizontal", command=canvas.xview)


content_frame = ttk.Frame(canvas)
content_frame.bind("<Configure>", lambda e: canvas.configure(width=e.width, scrollregion=canvas.bbox("all")))

canvas.create_window((0, 0), window=content_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbarv.set)
canvas.configure(xscrollcommand=scrollbarh.set)
canvas.grid(row=0, column=0, sticky="nsew")
scrollbarv.grid(row=0, column=1, sticky="ns")
scrollbarh.grid(row=1, column=0, sticky="ew")

canvas.columnconfigure(0, weight=1)
content_frame.columnconfigure(0, weight=1)
content_frame.rowconfigure(0, weight=0)
content_frame.rowconfigure(1, weight=1)
frame.update()
instr_lbl = tk.Label(master=content_frame, text="Tool to create the metadata marc xml necessary for ingest of Digital Representations into Alma Digital.\n'Metadata Spreadsheet' should be an Alma export in the format specified in the documentation. 'Item(s) parent folder' should be a folder of items, each in a folder named with the barcode reference.", anchor="nw", justify='left')
instr_lbl.bind('<Configure>', lambda e: instr_lbl.config(wraplength=(frame.winfo_width())-20))
instr_lbl.grid(row=0, column= 0, sticky="ew", padx=5)


ssheet = From_Spreadsheet(master=content_frame)
ssheet.grid(row=1, column= 0, sticky="ew", padx=5)

def _on_mousewheel(event):
   canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

canvas.bind_all("<MouseWheel>", _on_mousewheel)

tk.mainloop()