import pandas as pd
import xlsxwriter
from tkinter import *
from tkinter import ttk

window=Tk()
#add title for the window
window.title("Canview Quota Generator V1")
window.geometry("700x700+100+200")

my_notebook = ttk.Notebook(window)
my_notebook.grid(row=0)

myFrame1 = Frame(my_notebook, width=700, height=700)
myFrame2 = Frame(my_notebook, width=700, height=700)
myFrame3 = Frame(my_notebook, width=700, height=700)
myFrame1.pack(fill="both",expand=1)
myFrame2.pack(fill="both",expand=1)
myFrame3.pack(fill="both",expand=1)

my_notebook.add(myFrame1, text="Main")
my_notebook.add(myFrame2, text="Age Settings")
my_notebook.add(myFrame3, text="Region Settings")

Label(myFrame1, text="Demographics Used: ").grid(row=1, sticky=W)
var1 = IntVar()
Checkbutton(myFrame1, text="Gender", variable=var1).grid(row=2, sticky=W)
var2 = IntVar()
Checkbutton(myFrame1, text="Age", variable=var2).grid(row=3, sticky=W)
var3 = IntVar()
Checkbutton(myFrame1, text="Region", variable=var3).grid(row=4, sticky=W)
var4 = IntVar()
Checkbutton(myFrame1, text="Interlock", variable=var4).grid(row=5, sticky=W)

Label(myFrame1, text="Gender Marker: ").grid(row=2,column=1)
Label(myFrame1, text="Age Marker: ").grid(row=3,column=1)
Label(myFrame1, text="Region Marker: ").grid(row=4,column=1)

e1 = Entry(myFrame1, width=10, borderwidth=3)
e2 = Entry(myFrame1, width=10, borderwidth=3)
e3 = Entry(myFrame1, width=10, borderwidth=3)

e1.grid(row=2,column=2)
e2.grid(row=3,column=2)
e3.grid(row=4,column=2)

e1.insert(0, "gen")
e2.insert(0, "age")
e3.insert(0, "reg")

#GEN TAB
def genSelected(event):
    for x in range(int(ageCombo.get())):
        gen_entry = Entry(myFrame2)
        gen_entry.grid(row=x, column=1, pady=20, padx=5)

#GENDER SETTINGS
GENOPTIONS = [
    "Standard Male/Female",
    "Male/Female W/GEN LF"
]

genCombo = ttk.Combobox(myFrame1,value=GENOPTIONS)
genCombo.current(0)
genCombo.bind("<<ComboboxSelected>>", genSelected)
genCombo.grid(row=2,column=3, sticky=W)

#REG TAB
def regSelected(event):
    for x in range(int(ageCombo.get())):
        reg_entry = Entry(myFrame2)
        reg_entry.grid(row=x, column=1, pady=20, padx=5)
#REGION SETTINGS
REGOPTIONS = [
    "INCL EN QC + FR QC",
    "EXCL FR QC",
    "EXCL EN QC",
    "EXCL QC",
    "USA REGION",
    "Custom"
]

regCombo = ttk.Combobox(myFrame1,value=REGOPTIONS)
regCombo.current(0)
regCombo.bind("<<ComboboxSelected>>", regSelected)
regCombo.grid(row=4,column=3, sticky=W)

#AGE TAB
def ageSelected(event):
    for x in range(int(ageCombo.get())):
        age_entry = Entry(myFrame2)
        age_entry.grid(row=x, column=1, pady=20, padx=5)

#AGE SETTINGS MAIN
AGEOPTIONS = [
    "0",
    "2",
    "3",
    "4",
    "5",
    "6",
    "7",
    "8",
    "9",
    "10",
    "11",
    "12",
    "13",
    "14",
    "15"                       
]

ageCombo = ttk.Combobox(myFrame1,value=AGEOPTIONS)
ageCombo.current(0)
ageCombo.bind("<<ComboboxSelected>>", ageSelected)
ageCombo.grid(row=3,column=3, sticky=W)


#myFrame2.entry_widgets = [myFrame2.create_entry_widget(x) for x in range(myFrame2.n)]

#def create_entry_widget(myFrame2, x):
#    new_widget = Entry(myFrame2.master)
#    new_widget.pack()
#    new_widget.insert(0, x)
#    return new_widget

#CREATING XLS
def generateForm():
    workbook = xlsxwriter.Workbook('quota.xls')
    worksheet = workbook.add_worksheet('Defines')

    #sampleSize = input("What is the sample size of this study including any oversample: ")
    #cv_pid = input("Please type in the CV project ID (I.e., T001): ")

    # The workbook object is then used to add new 
    # worksheet via the add_worksheet() method.
    
    # Use the worksheet object to write
    # data via the write() method.
    worksheet.write('A1', 'Total')
    worksheet.write('B1', '1')
    if (var1.get() == 1):
        worksheet.write('A2', e1.get())
    if (var2.get() == 1):        
        worksheet.write('A3', e2.get())
    if (var3.get() == 1):
        worksheet.write('A4', e3.get())
    
    # Finally, close the Excel file
    # via the close() method.
    workbook.close()



mySubmit = Button(myFrame1, text="Generate Form", command=generateForm)
mySubmit.grid(row=50, sticky=S, column=2)

window.mainloop()