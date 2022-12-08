import tkinter as tk
import os
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog as fd
from pathlib import Path
from PyPDF2 import PdfFileWriter, PdfFileReader
from docx2pdf import convert
import win32gui, win32con

hide = win32gui.GetForegroundWindow()
win32gui.ShowWindow(hide , win32con.SW_HIDE)

def word2pdf_2(filedoc):
    convert(filedoc)

def reset_eof_of_pdf_return_stream(pdf_stream_in:list,):
    # find the line position of the EOF
    actual_line = 0
    for i, x in enumerate(pdf_stream_in[:-1]):
        if b'%%EOF' in x:
            actual_line = len(pdf_stream_in)-i
            print(f'EOF found at line position {-i} = actual {actual_line}, with value {x}')
            break

    # return the list up to that point
    return pdf_stream_in[actual_line]

def splitpdf(pdf_file_path, output_folder_path, InputText):
    pdf_file_path = pdf_file_path.replace('.docx', '')
    pdf_file_path = pdf_file_path+".pdf"
    print(pdf_file_path)
    # opens the file for reading

    pdf = PdfFileReader(pdf_file_path, strict=False)

    for page_num in range(pdf.numPages):
        pdfWriter = PdfFileWriter()
        pdfWriter.addPage(pdf.getPage(page_num))

        with open(os.path.join(output_folder_path, '{0}_{1}.pdf'.format(InputText, page_num+1)), 'wb') as f:
            pdfWriter.write(f)
            f.close()

# create the root window
root = tk.Tk()
root.title('PDFR')
root.resizable(False, False)
root.geometry('350x200')


def select_file(entry, label):
    txt="Espere mientras se compilan los archivos...";
    label.config(text = txt)
    
    InputText = "Documento"
    if (entry.get() != "" or entry.get() != " " or entry.get() != "  "):
        InputText = entry.get()
    
    print(InputText)
    
    
    '''try:'''
    filetypes = (
        ('Word Document', '*.docx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Seleccione su documento de salida',
        filetypes=filetypes)
    
    folder_selected = fd.askdirectory()

    word2pdf_2(filename)

    splitpdf(filename, folder_selected, InputText)

    txt="LOS DOCUMENTOS HAN SIDO GENERADOS CON ÉXITO"
    label.config(text = txt)
    messagebox.showinfo(message="LOS DOCUMENTOS HAN SIDO GENERADOS CON ÉXITO", title="Operación completa")

    '''except:
        txt="No se pudieron cargar los archivos"
        label.config(text = txt)'''

label = tk.Label( root, fg='#f00', text="")
label.place(x=20, y=180)

ttk.Label(text="Nombre de la plantilla: ").place(x=40, y=50)
entry = ttk.Entry()
entry.place(x=180, y=50)

open_button = ttk.Button(
    root,
    text='SELECCIONE UN ARCHIVO Y UN DIRECTORIO',
    command=lambda: select_file(entry, label)
)


open_button.place(x=50, y=90)


# run the application
root.mainloop()