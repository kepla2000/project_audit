from tkinter import *
from tkinter import filedialog
import os
import openpyxl

#Main Window Creation

root = Tk()
root.title("Audit X")
root.geometry("400x400")
#function for analysis and report of an xlsx file

def report_xlsx():
	pass

#function for analysis and report of a pdf file

def report_pdf():
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

root_filename = ""
def initial_button_function():
	root.filename = filedialog.askopenfilename(initialdir="C", title="Select file to Audit", filetypes=(("xlsx file", "*.png"),("pdf file", "*.pdf"),("docx file", "*.docx")))
	root_filename = root.filename
	Label12 = Label(root, text="You have selected" + root_filename)
	Label12.pack()
	return NONE

		
#detection of the extension name of the file

def extention_detection():
	modified_filename = f'r"{root_filename}"'
	ext = root_filename.split(".")[1]
	if ext == "xlsx":
		report_xlsx()
	elif ext == "pdf":
		report_pdf()
	elif ext == "docx":
		report_docx()

	pass

#Label for the first entry

Beginning = Label(root, text="Select file")
Beginning.pack()


#creating a button for browse

button_on_initial = Button(root, text="Browse" , command=initial_button_function)
button_on_initial.pack()



button_on_audit = Button(root, text="Audit", command=extention_detection)
button_on_audit.pack()

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




