# Develop by : P-HEX
# PDFIIDOCX Beta version v.1.0
# Created : 09/10/2021 (DD/MM/YYYY)

from pdf2docx import Converter
from docx import Document
from tkinter.filedialog import *
import threading
import pyglet

def startConvert(pdf_file, convertBtn, pdfFileLabel, findPdfBtn):

    top = Toplevel()

    threading.Thread(target=convert, args=(pdf_file, top, pdfFileLabel, findPdfBtn)).start()

    top.title('Processing')
    top.geometry(f"300x100+{root.winfo_x()}+{root.winfo_y()}")
    Message(top, text='Converting...', padx=100, pady=80, font=('PrintAble4U',15), bg=bgColor, fg='white').pack()

    convertBtn.config(state=DISABLED)
    convertBtn.update_idletasks()

    findPdfBtn.config(state=DISABLED)
    findPdfBtn.update_idletasks()

def convert(pdf_file, top, pdfFileLabel, findPdfBtn):

    file_path = os.path.basename(pdf_file)
    file_name = file_path.split('.pdf')

    cv = Converter(pdf_file)

    docx = createDocx(file_name[0])

    cv.convert(docx)
    cv.close()
    top.after(0, top.destroy)

    successMessageBox()

    findPdfBtn.config(state=NORMAL)
    pdfFileLabel.config(text='File is empty, Please select a PDF file', fg='red')

    findPdfBtn.update_idletasks()
    pdfFileLabel.update_idletasks()

def createDocx(pdfName):

    document = Document()
    docxFile = 'Output Files/' + str(pdfName) + '.docx'
    document.save(docxFile)

    return docxFile

def findPdf(convertBtn):

    fileOpen = askopenfilename(filetypes=[('PDF files', '*.pdf')])

    if fileOpen != '':
        pdfFileLabel.config(text=fileOpen, fg='#33FF00')

        convertBtn.config(state = NORMAL)
        convertBtn.update_idletasks()

def findOutput():

    path = os.path.realpath('Output Files')
    os.startfile(path)

def successMessageBox():

    top = Toplevel()
    top.title('Completed')
    top.geometry(f"300x100+{root.winfo_x()}+{root.winfo_y()}")
    Message(top, text='Converted.', padx=100, pady=80, font=('PrintAble4U', 15), bg=bgColor,fg='white').pack()
    top.after(3000, top.destroy)

if __name__=='__main__':

   bgColor = '#303030'
   btnColor = '#3ADDFF'

   pyglet.font.add_file('Font/PrintAble4U.ttf')

   root = Tk()
   root.iconbitmap('Icon/PDFIIDOCX.ico')
   root.configure(bg=bgColor)
   root.title('PDFIIDOCX (BETA VERSION V.1.0)')
   root.geometry('600x250')
   root.eval('tk::PlaceWindow . center')
   headLabel = Label(text='Convert PDF to DOCX', bg=bgColor, fg='white', font=('PrintAble4U',20)).place(x=200, y=20)

   pdfLabel = Label(text='PDF File : ', bg=bgColor, fg='white', font=('PrintAble4U', 15)).place(x=50, y=60)
   pdfFileLabel = Label(root, text='File is empty, Please select a PDF file', bg=bgColor, fg='red', font=('PrintAble4U', 15))
   pdfFileLabel.place(x=130, y=60)

   findPdfBtn = Button(root, text='Browse', command=lambda: findPdf(convertBtn), bg=btnColor, fg='black',font=('PrintAble4U', 15))
   findPdfBtn.place(x=500, y=60)

   convertBtn = Button(root, text='Convert', command=lambda: startConvert(pdfFileLabel.cget("text"), convertBtn, pdfFileLabel, findPdfBtn), font=('PrintAble4U', 15), bg=btnColor, fg='black', state=DISABLED)
   convertBtn.place(x=275, y=200)

   outputLabel = Label(text='Output folder : ', bg=bgColor, fg='white', font=('PrintAble4U',15)).place(x=50, y=100)
   findOutputBtn = Button(root, text='Open folder', command=findOutput, bg=btnColor, fg='black', font=('PrintAble4U',15)).place(x=170, y=100)

   devLabel = Label(root, text='Develop by : P-HEX', bg=bgColor, fg='#7500FF', font=('PrintAble4U', 13)).place(x=460, y=220)

   root.mainloop()