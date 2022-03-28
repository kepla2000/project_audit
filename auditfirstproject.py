from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import os
import sqlite3
import pandas as pd
from sqlalchemy import create_engine


#Main Window Creation

root = Tk()
root.title("Audit X")
root.geometry("400x400")

#function for analysis and report of an xlsx file

def report_xlsx():
	#transferring data from summarized results to database(semester 1 first file)
	df = pd.read_excel(root_filename ,sheet_name = 0 ,header=10, usecols='C:T', skiprows = 0)
	engine = create_engine('sqlite:///semester1_simple_com.sqlite')
	df.to_sql('semester1simple', con = engine, if_exists = 'replace', index = FALSE)

	#transferring data from complex results to database(semester 1 first file)
	df1 = pd.read_excel(root_filename , sheet_name = 1, usecols = 'C:AA', header = 13)
	df1.to_sql('semester1com1', con = engine, if_exists ='replace', index = FALSE)

	#transferring data from simple file to database (semester 2 first file)
	df2 = pd.read_excel(root_filename1 , sheet_name = 0, usecols = 'C:T', skiprows = 0, header = 10)
	df2.to_sql('semester2simple', con = engine, if_exists = 'replace', index = FALSE)


	#transferring data to a sql database complex file(semester 2 second file)  
	df3 = pd.read_excel(root_filename1, sheet_name = 1, usecols = 'C:AA', header = 13)
	df3.to_sql('semester2com1', con = engine, if_exists = 'replace', index = FALSE)

	#connecting program to database for value manipulation

	connection = sqlite3.connect('semester1_simple_com.sqlite')
	c = connection.cursor()

	#Error correction






#function for analysis and report of a pdf file

def report_pdf():
	#converting pdf file to an xlsx file 

	pass


#function for analysis and report of a docx file

def report_docx():
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
	global root_filename1
	root_filename1 = filedialog.askopenfilename(initialdir="C", title="Select file to Audit", filetypes=(("xlsx file", "*.xlsx"),("pdf file", "*.pdf"),("docx file", "*.docx"),)) 
	Label22.configure(text="File Selected: "+root_filename1)
	

 

def initial_button_function():
	global root_filename
	root_filename = filedialog.askopenfilename(initialdir="C", title="Select file to Audit", filetypes=(("xlsx file", "*.xlsx"),("pdf file", "*.pdf"),("docx file", "*.docx"),))
	Label11.configure(text="File Selected: "+root_filename)
	

		
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


#Label for the first entry,

Beginning = Label(root, text="Select file first")
Beginning.grid(row=1, column=3)

#creating a button for browse

button_on_initial = Button(root, text="Browse" , command=initial_button_function)
button_on_initial.grid(row=2 , column=2, padx=15)

Label11 = Label(root, text="No file selected")
Label11.grid(row=2, column=4)

Beginning = Label(root, text="Select second file")
Beginning.grid(row=3, column=3)

button_on_initial1 = Button(root, text="Browse", command=initial_button_function1)
button_on_initial1.grid(row=4, column=2, padx=15)


Label22 = Label(root, text="No file selected")
Label22.grid(row=4 , column=4)


button_on_audit = Button(root, text="Audit", command=extention_detection)
button_on_audit.grid(row=5, column=3)

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
edit_menu.add_command(labe="def14", command=command_edit3)


root.mainloop()





