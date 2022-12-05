import tkinter as tk
import os
from tkinter import ttk
from tkinter import filedialog as fd
from pathlib import Path
from PyPDF2 import PdfFileWriter, PdfFileReader
from docx2pdf import convert

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

def splitpdf(pdf_file_path, output_folder_path):
    pdf_file_path = pdf_file_path.replace('.docx', '')
    pdf_file_path = pdf_file_path+".pdf"
    print(pdf_file_path)
    # opens the file for reading

    pdf = PdfFileReader(pdf_file_path, strict=False)

    for page_num in range(pdf.numPages):
        pdfWriter = PdfFileWriter()
        pdfWriter.addPage(pdf.getPage(page_num))

        with open(os.path.join(output_folder_path, '{0}_Numero_{1}.pdf'.format("Doc", page_num+1)), 'wb') as f:
            pdfWriter.write(f)
            f.close()

# create the root window
root = tk.Tk()
root.title('PDFR')
root.resizable(False, False)
root.geometry('350x150')


def select_file():
    txt="Espere mientras se compilan los archivos...";
    label = tk.Label( root, fg='#f00', text=txt)
    label.pack(expand=True)
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

    splitpdf(filename, folder_selected)

    txt="LOS DOCUMENTOS HAN SIDO GENERADOS CON ÉXITO"
    label.config(text = txt)
    '''except:
        txt="No se pudieron cargar los archivos"
        label.config(text = txt)'''


# open button
open_button = ttk.Button(
    root,
    text='SELECCIONE UN ARCHIVO Y UN DIRECTORIO',
    command=select_file
)


open_button.pack(expand=True)


# run the application
root.mainloop()