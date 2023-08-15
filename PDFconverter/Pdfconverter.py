from tkinter import *
import PIL
from PIL import Image, ImageTk
from pdf2docx import Converter
from docx2pdf import convert
from fpdf import FPDF
import PyPDF2
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from pdf2image import convert_from_path
import img2pdf
import tkinter as tk
from tkinter import filedialog
from tkinter.messagebox import *
import tkinter.font as font
import os
import glob


def browsefiles():
	global filename
	filename = filedialog.askopenfilename(initialdir = "/",title = "Select a File",filetypes = (("all files","*.*"),("PDF files","*.pdf*"),("Word files","*.docx*"),("Word 2019 or below","*.doc*")))
	
	if filename.endswith('.pdf') or filename.endswith('.docx') or filename.endswith('.doc') or filename.endswith('.txt') or filename.endswith('.jpg') or filename.endswith('.png') or filename.endswith('.jpeg'):
		main_window_label_dyn.configure(text= "<< File Selected >> \n "+ filename,fg = 'white')
	else:
		showerror("Error","please provide a valid file \n\nOR \n\nDouble check your file's extension")

def pdf2docx():
	try:
		if filename.endswith(".pdf"):
			docx = filename + ".docx"
			pdf_file = filename
			docx_file = pdf_file.replace(".pdf",".docx")
		# convert pdf to docx
			cv = Converter(pdf_file)
			cv.convert(docx_file, start=0, end=None)
			cv.close()
			main_window_label_dyn.configure(text= "File converted!",fg = 'white')
		else:
			showinfo("PdfToWord","Please, provide a PDF File ")
	except NameError:
			showinfo("PdfToWord","Please, provide a PDF File ")


def pdf2txt():
	try:
		if filename.endswith(".pdf"):
			pdffileobj=open(filename,'rb')
			pdfreader=PyPDF2.PdfFileReader(pdffileobj)
			x=pdfreader.numPages
			pageobj=pdfreader.getPage(x-1)
			text = pageobj.extractText()
			file1=open(filename.replace(".pdf",".txt"),"a")
			file1.writelines(text)
			file1.close()
			main_window_label_dyn.configure(text ="File converted ")
		else:
			showinfo("Please","Choose a pdf File")
	except FileNotFoundError:
			showinfo("TextConversion","Please choose a pdf file")
	except NameError:
			showinfo("TextConversion","Please choose a pdf file")

def pdf2jpg():
	try:
		if filename.endswith(".pdf"):
			file = filename
			pages = convert_from_path(file)
			img = file.replace(".pdf","")
			count = 0
			for page in pages:
				count += 1
				jpeg_file = file.replace(".pdf","") + "-" + str(count)+".jpeg"
				page.save(jpeg_file.replace(".pdf","") ,'JPEG')
			main_window_label_dyn.configure(text = jpeg_file + "Created! ",fg = 'white')
		else:
			showinfo ("JpgConverter","Please choose a pdf file")	
	except NameError:
		showinfo ("JpgConverter","Please choose a pdf file")

def docx2pdf():
	try:
		if filename.endswith(".docx") or filename.endswith(".doc"):
			convert(filename,filename.replace(".docx",".pdf"))
			main_window_label_dyn.configure(text="File converted")
		else:
			showinfo ("WordToPdf","Please choose a pdf file")
	except NameError:
		showinfo("WordToPdf","Please Choose a Text File")

def text2pdf():
	try:
		pdf = FPDF()
		pdf.add_page()
		pdf.set_font("Arial", size = 15)
		f = open(filename,"r")
		for x in f:
			pdf.cell(200, 10, txt = x, ln = 1, align ='C')
		pdf.output(filename.replace(".txt",".pdf"))
		main_window_label_dyn.configure(text = "Converted! ",fg = 'white')
	except NameError:
		showinfo("ImageConverter","Please Choose a Text File")

def img2pdf():
	try:
		image1 = Image.open(filename)
		im1 = image1.convert('RGB')
		if filename.endswith('.jpeg'):
			im1.save(filename.replace(".jpeg",".pdf"))
			main_window_label_dyn.configure(text = "Converted! ",fg = 'white')
		if filename.endswith('.jpg'):
			im1.save(filename.replace(".jpg",".pdf"))
			main_window_label_dyn.configure(text = "Converted! ",fg = 'white')
		elif filename.endswith('.png'):
			im1.save(filename.replace(".png",".pdf"))
			main_window_label_dyn.configure(text = "Converted! ",fg = 'white')
		else:
			showinfo("ImgConverter","not a valid image format\n Only Jpeg & Png Extensions supported yet")
	except PIL.UnidentifiedImageError as e:
		showerror("ImgConverter","File provided is not a Image \n Only Jpeg & Png Extensions supported yet")
	except NameError:
		showinfo("ImageConverter","Please Choose a image")
	
	
def merge():
	dir = os.listdir("to-be-merged/")
	location = os.getcwd()
	if len(dir) <= 1:
		showinfo("MergePdf","Not enough files to merge\n Please add your files in \n -----[ to-be-merged ]----- \nfolder")
	else:
		pdfs = glob.glob("to-be-merged/*.pdf")
		merged_pdf = "Output/merged.pdf"
		merge_pdfs = PyPDF2.PdfFileMerger()

		for pdf in pdfs:
			merge_pdfs.append(open(pdf,'rb'))
		merge_pdfs.write(open(merged_pdf,'wb'))
		main_window_label_dyn.configure(text = "merged! ",fg = 'white')

def encrypt():
	print(filename)
	out = PdfFileWriter()
	cnt= PdfFileReader(filename)
	num = cnt.numPages

	for idx in range(num):
		page = cnt.getPage(idx)
		out.addPage(page)
		
	password = str(ent_pass.get())
	out.encrypt(password)
	with open(filename+"_encrypted.pdf", "wb") as f:
		out.write(f)
	ent_pass.delete(0,END)
	password_window.withdraw()
	main_window_label_dyn.configure(text = "Pdf Encrypted! \n Always remember your password :)",fg = 'white') 

def pswd():
	try:
		var = bool(filename)
		if var == TRUE:
			password_window.deiconify()
		else:
			showinfo("Please","please choose a file")
	except NameError as e:
		showinfo("Encryption Error","Please choose a file")
		
def back():
	ent_pass.delete(0,END)
	password_window.withdraw()
	

#fonts
fnt1 = ('calibri', 20)
fnt2 = ('consolas', 16, 'bold', 'underline')
fnt3 = ("times new roman", "32", "bold")
fnt4 = ("Freestyle Script", "72", "bold")

main_window = Tk()
main_window.geometry(f"{main_window.winfo_screenwidth()}x{main_window.winfo_screenheight()}+0+0")
main_window.maxsize(1920,1080)
main_window.minsize(1090,700)
main_window.title("PDF Converter")
main_window.iconbitmap("images\icon.ico")
main_window.configure(bg='#d9d9d9')


#list for images
bg=[]
logo=[]

#background image
variable1= Image.open("Images\gradient.png")
variable1=variable1.resize((1920,410), )
bg.append(ImageTk.PhotoImage(variable1))

#logos for buttons
logosf = Image.open("Images\select2.png")
logosf = logosf.resize((50,50), )
logo.append(ImageTk.PhotoImage(logosf))

logo1 = PhotoImage(file="Images\\u1.png")
logo2 = PhotoImage(file="Images\\u2.png")
logo3 = PhotoImage(file="Images\\u3.png")
logo4 = PhotoImage(file="Images\\u4.png")
logo5 = PhotoImage(file="Images\\u5.png") 
logo6 = PhotoImage(file="Images\\u6.png")
logo7 = PhotoImage(file="Images\\u7.png")
logo8 = PhotoImage(file="Images\\u8.png")
logo8 = logo8.zoom(1)

#frame for title
f1=Frame(main_window, height=150,bg='#d9d9d9' )
f1.pack(fill=X)

main_window_label_dyn = Label(f1, image=bg, text="Every tool you need to work with PDFs \nin one place", compound="center",fg="white",font=fnt4)
main_window_label_dyn.pack(side="top",pady=0)

#frame for buttons
f2=Frame(f1,height=468,bg='#d9d9d9')
f2.pack(fill=BOTH, padx=30)

f5=Frame(f1,height=468,bg='#d9d9d9')
f5.pack(fill=X, padx=30)

f3=Frame(f1,height=468,bg='#d9d9d9')
f3.pack(fill=X)

#frame for password

#4 buttons
btnwidth = 460
btnheight = 180

#5 buttons
#btnwidth = 380
#btnheight = 180

#3 buttons
#btnwidth = 640
#btnheight = 180

#Buttons
b6=Button(f2,image=logo,text = "Select File" , font=fnt2, bg='#d9d9d9', relief="groove", compound="left",borderwidth = 0, command= browsefiles,foreground = "green").pack(padx=800, pady=10)

b1=Button(f2, image=logo1,text="\n PDF to Word ", font=fnt3, width=btnwidth, height=btnheight, bg="white", relief="groove", compound="left", command = pdf2docx).pack(side="left",pady=3)
b2=Button(f2, image=logo2,text="\n PDF to Image ", font=fnt3, width=btnwidth, height=btnheight, bg="white", relief="groove", compound="left",command= pdf2jpg).pack(side="left",pady=3)
b3=Button(f2, image=logo3,text="\n PDF to Text ", font=fnt3, width=btnwidth, height=btnheight, bg="white", relief="groove", compound="left",command = pdf2txt).pack(side="left",pady=3)
b4=Button(f2, image=logo5,text="\n Merge Pdfs ", font=fnt3, width=btnwidth,height = btnheight, bg="white", relief="groove",compound="left",command = merge).pack(side="left",pady=3)
#b5=Button(f2, image=logo3,text="\n Word to PDF ", font=fnt3, width=btnwidth, height=btnheight, bg="white", relief="groove", compound="left").pack(side="left",pady=3)

b9=Button(f5, image=logo4,text="\n Word to PDF ", font=fnt3, width=btnwidth, height=btnheight, bg="white", relief="groove", compound="left", command = docx2pdf).pack(side="left",pady=3)
b10=Button(f5, image=logo6,text="\n Image to PDF ", font=fnt3, width=btnwidth, height=btnheight, bg="white", relief="groove", compound="left",command= img2pdf).pack(side="left",pady=3)
b11=Button(f5, image=logo8,text="\n Text to PDF ", font=fnt3, width=btnwidth, height=btnheight, bg="white", relief="groove", compound="left",command = text2pdf).pack(side="left",pady=3)
b12=Button(f5, image=logo7,text="\n Encrypt PDF ", font=fnt3, width=btnwidth, height=btnheight, bg="white", relief="groove", compound="left",command= pswd).pack(side="left",pady=3)
#b13=Button(f5, image=logo3,text="\n Word to PDF ", font=fnt3, width=btnwidth, height=btnheight, bg="white", relief="groove", compound="left").pack(side="left",pady=3)


b7=Button(f3, text="Exit ", font=fnt3, bg="white", relief="groove", compound="top", width=10,  command= exit).pack( pady=10)

password_window = Toplevel(main_window)
password_window.title("Encrypt")
password_window.geometry("500x200+600+400")
password_window.configure(bg='#d9d9d9')
ent_pass = Entry(password_window,borderwidth=4,width= 25,font = ('consolas', 16),show="*")
ent_pass.pack(pady=10)
enc_btn = Button(password_window,text = "Encrypt", width =20,height= 1,relief="groove",command = encrypt).pack(side="top",pady=2)
back_btn = Button(password_window,text = "Back", width =10,height= 1,relief="groove",command = back).pack(side="top",pady=2)

password_window.withdraw()
main_window.mainloop()