from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import os
import sqlite3
import pandas as pd
from sqlalchemy import create_engine
import tabula
from PyPDF2 import PdfFileReader 
import shutil
from tkinter import ttk
from docx2pdf import convert





#Main Window Creation

root = Tk()
root.title("Audit X")
root.geometry("800x300")



#function for analysis and report of an xlsx file

def report_xlsx():
	#transferring data from summarized results to database(semester 1 first file)
	df = pd.read_excel(root_filename ,sheet_name = 0 ,header=10, usecols='C:T', skiprows = 0)
	engine = create_engine('sqlite:///db.sqlite')
	df.to_sql('db', con = engine, if_exists = 'replace', index = FALSE)

	#transferring data from complex results to database(semester 1 first file)
	df1 = pd.read_excel(root_filename , sheet_name = 1, usecols = 'C:AA', header = 13)
	df1.to_sql('semester1com1', con = engine, if_exists ='replace', index = FALSE)

	#transferring data from simple file to database (semester 2 first file)
	df2 = pd.read_excel(root_filename1 , sheet_name = 0, usecols = 'C:T', skiprows = 0, header = 10)
	df2.to_sql('semester2simple', con = engine, if_exists = 'replace', index = FALSE)


	#transferring data to a sql database complex file(semester 2 second file)  
	df3 = pd.read_excel(root_filename1, sheet_name = 1, usecols = 'C:AA', header = 13)
	df3.to_sql('semester2com1', con = engine, if_exists = 'replace', index = FALSE)

	#transferring data to a sql database simple(supplimentary semester first file)
	df4 = pd.read_excel(root_filename2, sheet_name = 0, usecols = 'C:T', skiprows = 0, header = 10)
	df4.to_sql('semester3simple', con = engine, if_exists ='replace', index = FALSE)

	df5 = pd.read_excel(root_filename2 , sheet_name = 1, usecols = 'C:AA', header = 13)
	df5.to_sql('semester3com1' ,con = engine, if_exists ='replace' ,index = FALSE)


	#connecting program to database for value manipulation

	connection = sqlite3.connect('semester1_simple_com.sqlite')
	c = connection.cursor()
	#finding maxmimum number of subjects in that particular year
	c.execute("SELECT * FROM semester1simple")
	c.fetchall()

	#Error correction



#function for analysis and report of a pdf file

def report_pdf():
	#analysing the first semester pdf file
	#findng the number of pages of the specified pdf file
	with open(root_filename, "rb") as pdf_file:
		pdf_reader = PdfFileReader(pdf_file)
		acc_page = pdf_reader.numPages
		#creating folder for complex semester one filse (after convertion to xlsx file)
	directory = ["C:\\Users\\kepla\\Downloads\\New_folder1","C:\\Users\\kepla\\Downloads\\New_folder1\\complex_pdf_semester1(converted)","C:\\Users\\kepla\\Downloads\\New_folder1\\simple_pdf_semester1(converted)"]
	for i in directory:
			if os.path.exists(i):
				shutil.rmtree(i)
				os.mkdir(i)
				pass
			else:
				os.mkdir(i)

		
	#iterating through all the pages of the pdf file and converting it into an xlsx file and separating each page 
	#range scans values from the specified value to the end value + 1 
	for page in range(1, acc_page + 1):
		#creating a dataframe to contain information of the pdf file
		df_11 = tabula.read_pdf(root_filename, pages = page, stream = True)[0]
		df_11.to_excel(f"C:\\Users\\kepla\\Downloads\\New_folder1\\converted_semester1_{page}.xlsx")

	#moving complex files to a different folder
	for i in range(2, acc_page + 1):
		shutil.move(f"C:\\Users\\kepla\\Downloads\\New_folder1\\converted_semester1_{i}.xlsx","C:\\Users\\kepla\\Downloads\\New_folder1\\complex_pdf_semester1(converted)")

	#moving simple file to simple_pdf_semester1(converted)
	shutil.move(f"C:\\Users\\kepla\\Downloads\\New_folder1\\converted_semester1_1.xlsx","C:\\Users\\kepla\\Downloads\\New_folder1\\simple_pdf_semester1(converted)")
	#combining all the xlsx file into one master file
	current_directory = os.listdir("C:\\Users\\kepla\\Downloads\\New_folder1\\complex_pdf_semester1(converted)")
	excel_list = []

	for files in current_directory:
		excel_list.append(pd.read_excel(f"C:\\Users\\kepla\\Downloads\\New_folder1\\complex_pdf_semester1(converted)\\{files}"))

	#creating an empty dataframe
	excel_merged = pd.DataFrame()
	
	for excel_file in excel_list:
		excel_merged = excel_merged.append(excel_file, ignore_index = True)

	excel_merged.to_excel("C:\\Users\\kepla\\Downloads\\New_folder1\\complex_pdf_semester1(converted)\\final_converted_complex_result1.xlsx")

	#taking second file and analysing
	#finding the number of pages of the file
	with open(root_filename1, "rb") as pdf_file1:
		pdf_reader1 = PdfFileReader(pdf_file1)
		acc_page1 = pdf_reader1.numPages
		#creating folder for complex semester one filse (after convertion to xlsx file)
	directory1 = ["C:\\Users\\kepla\\Downloads\\New_folder2","C:\\Users\\kepla\\Downloads\\New_folder2\\complex_pdf_semester2(converted)","C:\\Users\\kepla\\Downloads\\New_folder2\\simple_pdf_semester2(converted)"]
	for i in directory1:
			if os.path.exists(i):
				shutil.rmtree(i)
				os.mkdir(i)
				pass
			else:
				os.mkdir(i)

		
	#iterating through all the pages of the pdf file and converting it into an xlsx file and separating each page 
	#range scans values from the specified value to the end value + 1 
	for page1 in range(1, acc_page1 + 1):
		#creating a dataframe to contain information of the pdf file
		df_22 = tabula.read_pdf(root_filename1, pages = page1, stream = True)[0]
		df_22.to_excel(f"C:\\Users\\kepla\\Downloads\\New_folder2\\converted_semester2_{page1}.xlsx")

	#moving complex files to a different folder
	for i in range(2, acc_page1 + 1):
		shutil.move(f"C:\\Users\\kepla\\Downloads\\New_folder2\\converted_semester2_{i}.xlsx","C:\\Users\\kepla\\Downloads\\New_folder2\\complex_pdf_semester2(converted)")

	#moving simple file to simple_pdf_semester1(converted)
	shutil.move(f"C:\\Users\\kepla\\Downloads\\New_folder2\\converted_semester2_1.xlsx","C:\\Users\\kepla\\Downloads\\New_folder2\\simple_pdf_semester2(converted)")
	#combining all the xlsx file into one master file
	current_directory1 = os.listdir("C:\\Users\\kepla\\Downloads\\New_folder2\\complex_pdf_semester2(converted)")
	excel_list1 = []

	for files1 in current_directory1:
		excel_list1.append(pd.read_excel(f"C:\\Users\\kepla\\Downloads\\New_folder2\\complex_pdf_semester2(converted)\\{files1}"))

	#creating an empty dataframe
	excel_merged1 = pd.DataFrame()
	
	for excel_file1 in excel_list1:
		excel_merged1 = excel_merged1.append(excel_file1, ignore_index = True)

	excel_merged1.to_excel("C:\\Users\\kepla\\Downloads\\New_folder2\\complex_pdf_semester2(converted)\\final_converted_complex_result2.xlsx")

	#taking supplimentary file and analysing

	with open(root_filename2, "rb") as pdf_file2:
		pdf_reader2 = PdfFileReader(pdf_file2)
		acc_page2 = pdf_reader2.numPages
		#creating folder for complex semester one filse (after convertion to xlsx file)
	directory1 = ["C:\\Users\\kepla\\Downloads\\New_folder3","C:\\Users\\kepla\\Downloads\\New_folder3\\complex_pdf_semester3(converted)","C:\\Users\\kepla\\Downloads\\New_folder3\\simple_pdf_semester3(converted)"]
	for i in directory1:
			if os.path.exists(i):
				shutil.rmtree(i)
				os.mkdir(i)
				pass
			else:
				os.mkdir(i)

		
	#iterating through all the pages of the pdf file and converting it into an xlsx file and separating each page 
	#range scans values from the specified value to the end value + 1 
	for page2 in range(1, acc_page2 + 1):
		#creating a dataframe to contain information of the pdf file
		df_33 = tabula.read_pdf(root_filename1, pages = page2, stream = True)[0]
		df_33.to_excel(f"C:\\Users\\kepla\\Downloads\\New_folder3\\converted_semester3_{page2}.xlsx")

	#moving complex files to a different folder
	for i in range(2, acc_page2 + 1):
		shutil.move(f"C:\\Users\\kepla\\Downloads\\New_folder3\\converted_semester3_{i}.xlsx","C:\\Users\\kepla\\Downloads\\New_folder3\\complex_pdf_semester3(converted)")

	#moving simple file to simple_pdf_semester1(converted)
	shutil.move(f"C:\\Users\\kepla\\Downloads\\New_folder3\\converted_semester3_1.xlsx","C:\\Users\\kepla\\Downloads\\New_folder3\\simple_pdf_semester3(converted)")
	#combining all the xlsx file into one master file
	current_directory2 = os.listdir("C:\\Users\\kepla\\Downloads\\New_folder3\\complex_pdf_semester3(converted)")
	excel_list2 = []

	for files2 in current_directory2:
		excel_list1.append(pd.read_excel(f"C:\\Users\\kepla\\Downloads\\New_folder3\\complex_pdf_semester3(converted)\\{files2}"))

	#creating an empty dataframe
	excel_merged2 = pd.DataFrame()
	
	for excel_file2 in excel_list2:
		excel_merged2 = excel_merged1.append(excel_file2, ignore_index = True)

	excel_merged1.to_excel("C:\\Users\\kepla\\Downloads\\New_folder3\\complex_pdf_semester3(converted)\\final_converted_complex_result3.xlsx")
	


#function for analysis and report of a docx file

def report_docx():
	#converting docx file to pdf file 
	
	pass

#functions for the events in the file menu

def command_file():
	pass

def command_file1():
	pass

def command_file2():
	pass

def command_file3():
	pass

def command_file4():
	root.quit()

#function for the event in the edit menu

def command_edit():
	pass

def command_edit1():
	pass

def command_edit2():
	pass

def command_edit3():
	pass


def initial_button_function1():
	#prievous file selected
	global root_filename1
	root_filename1 = filedialog.askopenfilename(initialdir="C", title="Select file to Audit", filetypes=(("xlsx file", "*.xlsx"),("pdf file", "*.pdf"),("docx file", "*.docx"),)) 
	Label22.configure(text="File Selected: "+root_filename1)
	

def initial_button_function():
	#current file selected
	global root_filename
	root_filename = filedialog.askopenfilename(initialdir="C", title="Select file to Audit", filetypes=(("xlsx file", "*.xlsx"),("pdf file", "*.pdf"),("docx file", "*.docx"),))
	Label11.configure(text="File Selected: "+root_filename)

def initial_button_function2():
	#supplimentary file selected
	global root_filename2
	root_filename2 = filedialog.askopenfilename(initialdir="C", title="Select file to Audit", filetypes=(("xlsx file", "*.xlsx"),("pdf file", "*.pdf"),("docx file", "*.docx"),))
	Label33.configure(text="File Selected: "+root_filename2)
	
#detection of the extension name of the file

def extention_detection():
	ext = root_filename.split(".")[-1]
	ext1 = root_filename1.split(".")[-1]
	if ext and ext1 == "xlsx":
		report_xlsx()
	elif ext and ext1 == "docx":
		report_docx()
	elif ext and ext1 == "pdf":
		report_pdf()

def next():
	
	#checking for the options chosen (the year and the semester)
	options = clicked.get()
	options1 = clicked1.get()
	global Label11, Label22, Label33, button_on_initial1, button_on_audit, button_on_initial, button_on_audit, Beginning, Label33, Beginning2, button_back
	button_back = Button(frame, text="back" , command=destroyer)


	

	if (options == "year1") and (options1 == "semester1"):
		messagebox.showerror("Error", "Year1 semster1 cannot be compared with any other file")
	elif (options == "year1" or options == "year3" or options =="year2" or options =="year4") and (options1 == "semester2" or options1 == "semester3"):
		'''top = Toplevel()
		top.geometry("800x300")
		top.title("Audit")'''
		dropbox.destroy()
		dropbox1.destroy()
		button_next.destroy()
		button_back.grid(row=14, column=5, padx=15, pady=15)

	
		Beginning = Label(frame, text="Select current semester file")
		Beginning.grid(row=1, column=3)

		Label11 = Label(frame, text="No file selected")
		Label11.grid(row=2, column=4)

		button_on_initial = Button(frame, text="Browse" , command=initial_button_function)
		button_on_initial.grid(row=2 , column=1, padx=15)


		Beginning2 = Label(frame, text="Select previous semester file")
		Beginning2.grid(row=3, column=3)

		Label22 = Label(frame, text="No file selected")
		Label22.grid(row=4, column=4)

		button_on_initial1 = Button(frame, text="Browse", command=initial_button_function1)
		button_on_initial1.grid(row=4, column=1)


		button_on_audit = Button(frame, text="Audit", command=extention_detection)
		button_on_audit.grid(row=8, column=2,padx =15)
	
	elif (options == "year2" or options == "year3" or options == "year4") and (options1 == "semester1"):
		'''top1 = Toplevel()
		top1.geometry("800x300")
		top1.title("Audit")'''
		dropbox.destroy()
		dropbox1.destroy()
		button_next.destroy()
		button_back.grid(row=14, column=5, padx=15, pady=15)

		Beginning = Label(frame, text="Select current semester file")
		Beginning.grid(row=1, column=3)

		Label11 = Label(frame, text="No file selected")
		Label11.grid(row=2, column=4)

		button_on_initial = Button(frame, text="Browse", command=initial_button_function)
		button_on_initial.grid(row=2, column=1, padx=15)


		Beginning2 = Label(frame, text="Select previous semester file")
		Beginning2.grid(row=3, column=3)

		Label22 = Label(frame, text="No file selected")
		Label22.grid(row=4, column=4)

		button_on_initial1 = Button(frame, text="Browse", command=initial_button_function1)
		button_on_initial1.grid(row=4, column=1)


		Beginning2 = Label(frame, text="Select previous semester3 file")
		Beginning2.grid(row=5, column=3)

		Label33 = Label(frame, text="No file selected")
		Label33.grid(row=6, column=4)

		button_on_initial2 = Button(frame, text="Browse", command=initial_button_function2)
		button_on_initial2.grid(row=6, column=1)


		button_on_audit = Button(frame, text="Audit", command=extention_detection)
		button_on_audit.grid(row=8, column=2,padx =15)

		li = [Label11, Label22, Label33, button_on_initial1, button_on_audit, button_on_initial, button_on_audit, Beginning, Label33, Beginning2,button_on_initial2]

def destroyer():
	frame.destroy()
	back()


def back(): 

	global dropbox, dropbox1, clicked1, clicked, frame, button_next
	frame = LabelFrame(root, text="Please select current academic year and semester" , padx=5, pady=5)
	frame.pack(padx=20, pady=20)

	clicked = StringVar()
	clicked.set("year1")

	clicked1 = StringVar()
	clicked1.set("semester1")

	dropbox = OptionMenu(frame,clicked,"year1", "year2", "year3", "year4",)
	dropbox.grid(row=4, column=5, padx=155, pady=15)

	dropbox1 = OptionMenu(frame, clicked1, "semester1","semester2","semester3")
	dropbox1.grid(row=10, column=5, padx=15, pady=15)

	button_next = Button(frame, text="Next" , command=next)
	button_next.grid(row=14, column=5, padx=15, pady=15)

back()




















"""Beginning = Label(root, text="Select semester1 file")
Beginning.grid(row=1, column=3)

#creating a button for browse

button_on_initial = Button(root, text="Browse" , command=initial_button_function)
button_on_initial.grid(row=2 , column=2, padx=15)



Label11 = Label(root, text="No file selected")
Label11.grid(row=2, column=4)

Beginning = Label(root, text="Select semester2 file")
Beginning.grid(row=4, column=3)

button_on_initial1 = Button(root, text="Browse", command=initial_button_function1)
button_on_initial1.grid(row=6, column=2, padx=15)




Label22 = Label(root, text="No file selected")
Label22.grid(row=6 , column=4)


button_on_audit = Button(root, text="Audit", command=extention_detection)
button_on_audit.grid(row=9, column=3,padx =15)

button_on_supplimentary = Button(root, text="Browse", command=initial_button_function2)
button_on_supplimentary.grid(row=8, column=2, padx=15)



Label33 = Label(root, text="No file selected")
Label33.grid(row=8, column=4)

Label44 = Label(root, text = "Select supplimentary file")
Label44.grid(row=7, column=3)


#creating a menu

my_menu = Menu(root)
root.config(menu=my_menu)

#creating the cascading items for the file menu

file_menu = Menu(my_menu)
my_menu.add_cascade(label="File",menu=file_menu)
file_menu.add_command(label="New", command=command_file)
file_menu.add_command(label="def1", command=command_file1)
file_menu.add_command(label="def2", command=command_file2)
file_menu.add_command(label="def3", command=command_file3)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=command_file4)

#creating the cascading items for the edit menu 

edit_menu = Menu(my_menu)
my_menu.add_cascade(label="Edit", menu=edit_menu)
edit_menu.add_command(labe="def11", command=command_edit)
edit_menu.add_command(labe="def12", command=command_edit1)
edit_menu.add_command(labe="def13", command=command_edit2)
edit_menu.add_command(labe="def14", command=command_edit3)"""

root.mainloop()





