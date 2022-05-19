import enum
import pandas as pd
import xlsxwriter
import customtkinter as CustTK
import tkinter as tk
from tkinter import *
from tkinter import ttk

CustTK.set_appearance_mode("System")  # Modes: system (default), light, dark
CustTK.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green
window=CustTK.CTk()
#add title for the window
window.title("Canview Quota Generator V1")
window.geometry("1300x900+150+250")

my_notebook = ttk.Notebook(window)
my_notebook.grid(row=0)

myFrame1 = CustTK.CTkFrame(my_notebook, width=1300, height=850, padx=20, pady=20)
myFrame2 = CustTK.CTkFrame(my_notebook, width=1300, height=850, padx=20, pady=20)
myFrame3 = CustTK.CTkFrame(my_notebook, width=1300, height=850, padx=20, pady=20)
myFrame4 = CustTK.CTkFrame(my_notebook, width=1300, height=850, padx=20, pady=20)
myFrame1.pack(fill="both",expand=1)
myFrame2.pack(fill="both",expand=1)
myFrame3.pack(fill="both",expand=1)
myFrame4.pack(fill="both",expand=1)

#add tabs
my_notebook.add(myFrame1, text="Main")
my_notebook.add(myFrame2, text="Age Settings")
my_notebook.add(myFrame3, text="Region Settings")
my_notebook.add(myFrame4, text="Gender Settings")

#Main page - 1st row labels + N total
CustTK.CTkLabel(myFrame1, text="Hard Quota:").grid(row=1, sticky=W)
CustTK.CTkLabel(myFrame1, text="Demos Used:").grid(row=1, column=1, sticky=W)
CustTK.CTkLabel(myFrame1, text="Number of variable groups").grid(row=1, column=7, sticky=E)
CustTK.CTkLabel(myFrame1, text="TOTAL N: ").grid(row=1, column=4)
#Main page - 2nd set of interlock not sure if we should use
# var5 = IntVar()
# genButton2 = CustTK.CTkCheckBox(myFrame1, state=DISABLED, text="Gender", variable=var5)
# genButton2.grid(row=2, column=8, sticky=W)
# var6 = IntVar()
# ageButton2 = CustTK.CTkCheckBox(myFrame1, state=DISABLED, text="Age", variable=var6)
# ageButton2.grid(row=3, column=8, sticky=W)
# var7 = IntVar()
# regButton2 = CustTK.CTkCheckBox(myFrame1, state=DISABLED, text="Region", variable=var7)
# regButton2.grid(row=4, column=8, sticky=W)
# var8 = IntVar()
# intButton2 = CustTK.CTkCheckBox(myFrame1, state=DISABLED, text="Interlock 2", variable=var8)
# intButton2.grid(row=5, column=8, sticky=W)

is_on=True
#defined function to enable 2nd set of interlocks when 1st interlock button clicked
# def enableInt():
#     global is_on

#     if is_on:
#         genButton2.config(state=NORMAL)
#         ageButton2.config(state=NORMAL)
#         regButton2.config(state=NORMAL)
#         intButton2.config(state=NORMAL)       
#         is_on=False
#     else:
#         genButton2.deselect()
#         ageButton2.deselect()
#         regButton2.deselect()
#         intButton2.deselect()

#         genButton2.config(state=DISABLED)
#         ageButton2.config(state=DISABLED)
#         regButton2.config(state=DISABLED)
#         intButton2.config(state=DISABLED)
#         is_on=True       

def disableHQ():
    global is_on

    if is_on:
        genHQ.deselect()
        ageHQ.deselect()
        regHQ.deselect()
        genHQ.config(state=DISABLED)
        ageHQ.config(state=DISABLED)
        regHQ.config(state=DISABLED)
        is_on=False
    else:
        genHQ.config(state=NORMAL)
        ageHQ.config(state=NORMAL)
        regHQ.config(state=NORMAL)
        is_on=True          
#Main page - 1st set of interlock

var1 = IntVar()
genButton1 = CustTK.CTkCheckBox(myFrame1, text="Gender", variable=var1)
genButton1.grid(row=2, column=1, sticky=W)
var2 = IntVar()
ageButton1 = CustTK.CTkCheckBox(myFrame1, text="Age", variable=var2)
ageButton1.grid(row=3, column=1, sticky=W)
var3 = IntVar()
regButton1 = CustTK.CTkCheckBox(myFrame1, text="Region", variable=var3)
regButton1.grid(row=4, column=1, sticky=W)
var4 = IntVar()
interlockButton1 = CustTK.CTkCheckBox(myFrame1, text="Interlock 1", variable=var4)
interlockButton1.grid(row=5, column=1, sticky=W, pady=10)

#Main page - HQ checkboxes for 1st set of interlock
var11 = IntVar()
genHQ = CustTK.CTkCheckBox(myFrame1, variable=var11, text="")
genHQ.grid(row=2)
var12 = IntVar()
ageHQ = CustTK.CTkCheckBox(myFrame1, variable=var12, text="")
ageHQ.grid(row=3)
var13 = IntVar()
regHQ = CustTK.CTkCheckBox(myFrame1, variable=var13, text="")
regHQ.grid(row=4)
var14 = IntVar()
CustTK.CTkCheckBox(myFrame1, variable=var14, text="", command=disableHQ).grid(row=5)


ttk.Separator(myFrame1, orient='horizontal').grid(row=6, columnspan=200, pady = 20, sticky="ew")
ttk.Separator(myFrame1, orient='horizontal').grid(row=20, columnspan=200, pady = 20, sticky="ew")

CustTK.CTkLabel(myFrame1, text="Gender Marker: ").grid(row=2,column=2)
CustTK.CTkLabel(myFrame1, text="Age Marker: ").grid(row=3,column=2)
CustTK.CTkLabel(myFrame1, text="Region Marker: ").grid(row=4,column=2)

e1 = CustTK.CTkEntry(myFrame1, width=50, border_width=2, corner_radius=5)
e2 = CustTK.CTkEntry(myFrame1, width=50, border_width=2, corner_radius=5)
e3 = CustTK.CTkEntry(myFrame1, width=50, border_width=2, corner_radius=5)

e1.grid(row=2,column=3, pady=2)
e2.grid(row=3,column=3, pady=2)
e3.grid(row=4,column=3, pady=2)

e1.insert(0, "gen")
e2.insert(0, "age")
e3.insert(0, "reg")


CustTK.CTkLabel(myFrame1, text="Gender Var: ").grid(row=2,column=4)
CustTK.CTkLabel(myFrame1, text="Age Var: ").grid(row=3,column=4)
CustTK.CTkLabel(myFrame1, text="Region Var: ").grid(row=4,column=4)

#This function ensures only integers/digitscan be typed into sample size field
def validateInt10(P):
    if len(P)== 0 or len(P) < 10 and P.isdigit():
        return True
    else:
        return False
vcmdInt = (myFrame1.register(validateInt10),'%P')

#This function ensures only floats with leading 0s/periods can be typed into the percentage fields
def validateFloat(P):
    text = P  
    print('text:', text)

    parts = text.split('.')
    parts_number = len(parts)

    if parts_number > 2:
        #print('too many dots')
        return False

    if parts_number > 1 and parts[1]: # don't check empty string
        if not parts[1].isdecimal() or len(parts[1]) > 5:
            #print('wrong second part')
            return False

    if parts_number > 0 and parts[0]: # don't check empty string
        if not parts[0].isdecimal() or len(parts[0]) > 1 or ('0' not in parts[0]):
            #print('wrong first part')
            return False

    return True

e4 = CustTK.CTkEntry(myFrame1, width=125, border_width=2, corner_radius=5)
e5 = CustTK.CTkEntry(myFrame1, width=125, border_width=2, corner_radius=5)
e6 = CustTK.CTkEntry(myFrame1, width=125, border_width=2, corner_radius=5)
e10 = CustTK.CTkEntry(myFrame1, width=150, border_width=2, corner_radius=5, validate="key", validatecommand=(vcmdInt))

e10.grid(row=1,column=5)
e4.grid(row=2,column=5)
e5.grid(row=3,column=5)
e6.grid(row=4,column=5)

e4.insert(0, "GENDER")
e5.insert(0, "QUOTAGERANGE")
e6.insert(0, "REGION")

sv1 = StringVar()
sv2 = StringVar()
sv3 = StringVar()

e7 = CustTK.CTkEntry(myFrame1, width=30, border_width=2, corner_radius=5, textvariable=sv2)
e8 = CustTK.CTkEntry(myFrame1, width=30, border_width=2, corner_radius=5, textvariable=sv3)
e9 = CustTK.CTkEntry(myFrame1, width=30, border_width=2, corner_radius=5, textvariable=sv1)

e7.grid(row=2,column=7)
e8.grid(row=3,column=7)
e9.grid(row=4,column=7)

e7.insert(0, 0)
e8.insert(0, 0)
e9.insert(0, 0)

e7.configure(state='disabled')
e8.configure(state='disabled')
e9.configure(state='disabled')

#Begins some code to setup the CELL BALANCE FIELDS
#Main page - Checkbox buttons to indicate which demos shoudl be balanced for in cell balance

var10 = IntVar()
genBalance2 = CustTK.CTkCheckBox(myFrame1, text="Gender", variable=var10)
genBalance2.grid(row=7, column=2, sticky=W, pady=2)
var11 = IntVar()
ageBalance2 = CustTK.CTkCheckBox(myFrame1, text="Age", variable=var11)
ageBalance2.grid(row=8, column=2, sticky=W, pady=2)
var12 = IntVar()
regBalance2 = CustTK.CTkCheckBox(myFrame1, text="Region", variable=var12)
regBalance2.grid(row=9, column=2, sticky=W, pady=2)


#legacy vs cross table setup
cellBalType = IntVar(0)
def radiobutton_event():
    print("radiobutton toggled, current value:", cellBalType.get())
radiobutton_1 = CustTK.CTkRadioButton(myFrame1, text="Cross Table Method",command=radiobutton_event, variable= cellBalType, value=1)
radiobutton_2 = CustTK.CTkRadioButton(myFrame1, text="Legacy Method",command=radiobutton_event, variable= cellBalType, value=2)
radiobutton_1.grid(row=7, column=1, sticky=W, pady=2, padx=10)
radiobutton_2.grid(row=8, column=1, sticky=W, pady=2, padx=10)

#Entry fields for number of cells
CustTK.CTkLabel(myFrame1, text="Num Cells").grid(row=7, column=0)
e_cb1 = CustTK.CTkEntry(myFrame1, width=50, border_width=2, corner_radius=5, validate="key", validatecommand=(vcmdInt))
e_cb1.grid(row=8,column=0)
CustTK.CTkLabel(myFrame1, text="Num Picks").grid(row=9, column=0)
e_cb2 = CustTK.CTkEntry(myFrame1, width=50, border_width=2, corner_radius=5, validate="key", validatecommand=(vcmdInt))
e_cb2.grid(row=10,column=0)

#cell qualifications
CustTK.CTkLabel(myFrame1, text="Qual Variable, only one is needed/allowed (i.e., cellQual)").grid(row=7, column=4,columnspan=4)
e_cb3 = CustTK.CTkEntry(myFrame1, width=450, state=DISABLED, border_width=2, corner_radius=5)
e_cb3.grid(row=8,column=4,columnspan=4,padx=20)
CustTK.CTkLabel(myFrame1, text="Qual Row Labels, separated by commas (I.e., r1,r2,r3,r4)").grid(row=9, column=4,columnspan=4)
e_cb4 = CustTK.CTkEntry(myFrame1, width=450, state=DISABLED, border_width=2, corner_radius=5)
e_cb4.grid(row=10,column=4,columnspan=4,padx=20)
CustTK.CTkLabel(myFrame1, text="Priority Row Labels, separated by commas (I.e., r1,r2,r3,r4)").grid(row=11, column=4,columnspan=4)
e_cb5 = CustTK.CTkEntry(myFrame1, width=450, state=DISABLED, border_width=2, corner_radius=5)
e_cb5.grid(row=12,column=4,columnspan=4,padx=20)


#defined function to enable cell qualification fields
def enableCQ():
    global is_on

    if is_on:
        e_cb3.config(state=NORMAL)
        e_cb4.config(state=NORMAL)     
        is_on=False
    else:
        e_cb3.config(state=DISABLED)
        e_cb4.config(state=DISABLED)
        is_on=True       
#checkbox to indicate if cells are qualified for or if completely random
var13 = IntVar()
qualBalance = CustTK.CTkCheckBox(myFrame1, text="Cell Qualifications?", variable=var13, command=enableCQ)
qualBalance.grid(row=10, column=2, sticky=W, pady=2,columnspan=2)

#Arrays used to calculate hard quotas
InterlockArray = []
#genArray = []
#ageArray = []
#regArray = []

#GEN TAB
genDict = {
  1: 2,
  2: 4
}
genNames = {
  1: "Male",
  2: "Female",
  3: "Other 1",
  4: "Other 2",
  5: "Other 3",
  6: "Other 4"
}
genPerc = {
  1: .50,
  2: .50,
  3: 0,
  4: 0,
  5: 0,
  6: 0
}

sv2.trace("w", lambda name, index, mode, sv2=sv2: callback2(sv2))
    
def callback2(sv):
    # myFrame4.entries=[]
    # myFrame4.GenPercList = []

    #print (sv.get())

    for widget in myFrame4.winfo_children():
        widget.destroy()

    #created this array to be able to keep track of whatever is entered into the entries and recalculate dynamically
    myFrame4.genArray = []
    vcmdFt = (myFrame4.register(validateFloat),'%P')
    for y in range(int(sv.get())):
        myFrame4.gen_entry = CustTK.CTkEntry(myFrame4,validate="key",validatecommand=(vcmdFt))
        gen_name = CustTK.CTkEntry(myFrame4)
        
        #Create gender name text boxes
        gen_name.grid(row=y, column=2, pady=10, padx=5)
        #Fill gender name text boxes with default gender labels
        gen_name.insert(0, (genNames[y+1])) 
        #Create gender % text boxes
        myFrame4.gen_entry.grid(row=y, column=3, pady=10, padx=5)
        #Fill gender % text boxes with default gender %
        myFrame4.gen_entry.insert(0, (genPerc[y+1])) 
        myFrame4.genArray.append(myFrame4.gen_entry)
        Label(myFrame4, text="Gen" + str(y+1)).grid(row=y,column=1)
        myFrame4.entries.append(gen_name) 
        myFrame4.GenPercList.append(genPerc[y+1]) 

def genSelected(event):
    e7.configure(state='normal')
    #e7.delete(0,"end")

    for widget in myFrame4.winfo_children():
        widget.destroy()
    myFrame4.entries=[]
    myFrame4.GenPercList = []

    #e7.insert(0,(int(genDict[genCombo.current()])))
    sv2.set(int(genDict[genCombo.current()]))
    # for x in range(int(genDict[genCombo.current()])):
    #         gen_entry = Entry(myFrame4)
    #         Label(myFrame4, text="Gender " + str(x+1)).grid(row=x,column=1)
    #         gen_entry.grid(row=x, column=2, pady=10, padx=5)   
    #         gen_entry.insert(0, (genNames[x+1])) 
            #myFrame4.labels.append(gen_entry)     

    if genCombo.current() == 1:
        # e7.insert(0,(int(genDict[genCombo.current()])))
        e7.configure(state='disabled')        
        # for x in range(int(genDict[genCombo.current()])):
        #     gen_entry = Entry(myFrame4)
        #     Label(myFrame4, text="Gender " + str(x+1)).grid(row=x,column=1)
        #     gen_entry.grid(row=x, column=2, pady=10, padx=5)   
        #     gen_entry.insert(0, (genNames[x+1])) 
        #     myFrame4.labels.append(gen_entry)   
    # else: 
        # e7.insert(0,(int(genDict[genCombo.current()])))
        # for x in range(int(genDict[genCombo.current()])):
        #     gen_entry = Entry(myFrame4)
        #     Label(myFrame4, text="Gender " + str(x+1)).grid(row=x,column=1)
        #     gen_entry.grid(row=x, column=2, pady=10, padx=5)   
        #     gen_entry.insert(0, (genNames[x+1])) 
        #     myFrame4.labels.append(gen_entry)  
    
#GENDER SETTINGS
GENOPTIONS = [
    "SELECT ONE",
    "Standard Male/Female",
    "Male/Female W/GEN LF"
]

genCombo = ttk.Combobox(myFrame1,value=GENOPTIONS, justify=CENTER)
genCombo.current(0)
genCombo.bind("<<ComboboxSelected>>", genSelected)
genCombo.grid(row=2,column=6, sticky=W)

#REG TAB
regDict = {
  1: ["WEST","ONTARIO","QC","ATLANTIC"],
  2: ["WEST","ONTARIO","EN QC","ATLANTIC"],
  3: ["WEST","ONTARIO","FR QC","ATLANTIC"],
  4: ["WEST","ONTARIO","ATLANTIC"],
  5: ["BC","PRAIRIES","ONTARIO","EN QC","FR QC","ATLANTIC"],
  6: ["BC","PRAIRIES","ONTARIO","EN QC","ATLANTIC"],
  7: ["BC","PRAIRIES","ONTARIO","FR QC","ATLANTIC"],
  8: ["BC","PRAIRIES","ONTARIO","ATLANTIC"],
  9: ["NORTHEAST","MIDWEST","SOUTH","WEST"],
  10: [""]  
}
regCounter = {
  1: 4,
  2: 4,
  3: 4,
  4: 3,
  5: 6,
  6: 5,
  7: 5,
  8: 4,
  9: 4,
  10: [""]  
}
regPerc = {
  1: [0.3012,0.3898,0.2362,0.0728],
  2: [0.3850,0.4980,0.0240,0.0930],
  3: [0.3106,0.4020,0.2122,0.0752],
  4: [0.3942,0.5104,0.0954],
  5: [0.1323,0.1688,0.3898,0.024,0.2122,0.0729],
  6: [0.169,0.216,0.498,0.024,0.093],
  7: [0.1365,0.1741,0.402,0.2122,0.0752],
  8: [0.1732,0.221,0.5104,0.0954],
  9: [0.178186900785363,0.215534586540639,0.372431791932523,0.233846720741475],
  10: [0]
}

def callback1(sv):
    if (regCombo.current() == 10):
        myFrame3.entries=[]
        myFrame3.regArray = []

        for widget in myFrame3.winfo_children():
            widget.destroy()
        CustTK.CTkLabel(myFrame3, text="Region Label").grid(row=0,column=1)
        CustTK.CTkLabel(myFrame3, text="Region Percent").grid(row=0,column=2)    

        for y in range(int(sv.get())):
            vcmdFt = (myFrame3.register(validateFloat),'%P')
            myFrame3.reg_entry = CustTK.CTkEntry(myFrame3,validate="key",validatecommand=(vcmdFt))
            reg_name = CustTK.CTkEntry(myFrame3)
            
            reg_name.grid(row=y+1, column=1, pady=10, padx=5)
            myFrame3.reg_entry.grid(row=y+1, column=2, pady=10, padx=5)
            myFrame3.regArray.append(myFrame3.reg_entry)
            reg_name.insert(0, "REGION " + str(y+1))
            myFrame3.entries.append(reg_name) 

sv1.trace("w", lambda name, index, mode, sv=sv1: callback1(sv))

def regSelected(event):
    e9.configure(state='normal')
    #e9.insert(0, len(regDict[startIndex]))
    sv1.set(regCounter[regCombo.current()])
    #e9.insert(0, regCounter[startIndex])
    e9.configure(state='disabled')
    
    myFrame3.RegPercList=[]
    myFrame3.regArray = []

    for widget in myFrame3.winfo_children():
        widget.destroy()

    CustTK.CTkLabel(myFrame3, text="Region Label").grid(row=0,column=1)
    CustTK.CTkLabel(myFrame3, text="Region Percent").grid(row=0,column=2)

    if not(regCombo.current() == 10):
        startIndex = regCombo.current()
        
        for x in range(len(regDict[startIndex])):
            myFrame3.reg_entry = CustTK.CTkEntry(myFrame3)
            CustTK.CTkLabel(myFrame3, text=str((regDict[startIndex])[x])).grid(row=x+1,column=1)
            myFrame3.reg_entry.grid(row=x+1, column=2, pady=10, padx=5)
            myFrame3.reg_entry.insert(0, ((regPerc[startIndex])[x])) 
            myFrame3.regArray.append(myFrame3.reg_entry)
            myFrame3.RegPercList.append((regPerc[startIndex])[x])
    else:
        sv1.set(4)
        e9.configure(state='normal')

#REGION SETTINGS
REGOPTIONS = [
    "SELECT ONE",
    "INCL EN/FR QC",
    "EXCL FR QC",
    "EXCL EN QC",
    "EXCL QC",
    "INCL EN/FR QC (BC)",
    "EXCL FR QC (BC)",
    "EXCL EN QC (BC)",
    "EXCL QC (BC)",
    "USA REGION",
    "Custom"    
]

regCombo = ttk.Combobox(myFrame1,value=REGOPTIONS, justify=CENTER)
regCombo.current(0)
regCombo.bind("<<ComboboxSelected>>", regSelected)
regCombo.grid(row=4,column=6, sticky=W)
vAge = StringVar()
#AGE TAB
#Function for calculate age percentage button; calculates age % breakdowns based upon the entries
def calcAge(): 
    #print ("TEST FRAME 2 BUTTON")
    #print (int(myFrame2.lowRange[0].get()))
    #print (ageCensus[int(myFrame2.lowRange[0].get())])

    #myFrame2.totalList[0].insert(0, ageCensus[int(myFrame2.lowRange[0].get())])
    
    tempTotal = 0
    myFrame2.AgePercList = []

    #This loop will compile and insert the 'total population' within each age range and insert it into the total in range column
    for x in range(int(ageCombo.get())):
        tempCount = 0
        for y in range(int(myFrame2.lowRange[x].get()),int(myFrame2.highRange[x].get())+1):
            tempCount = tempCount + ageCensus[y]
        myFrame2.totalList[x].configure(state='normal')
        myFrame2.totalList[x].insert(0, tempCount)
        myFrame2.totalList[x].configure(state='disabled')
        tempTotal = tempTotal + int(myFrame2.totalList[x].get())
    
    #print (tempTotal)
    #This loop will compile and insert the 'percentage' of each age ranges' total population relative to the overall sum and insert it into the percentage column
    for j in range(int(ageCombo.get())):
        myFrame2.PercList[j].configure(state='normal')
        myFrame2.PercList[j].insert(0, (int(myFrame2.totalList[j].get())/tempTotal))
        myFrame2.PercList[j].configure(state='disabled')
        myFrame2.AgePercList.append(int(myFrame2.totalList[j].get())/tempTotal)
        
def validateAge(P):
    if len(P)== 0 or len(P) < 3 and P.isdigit():
        return True
    else:
        return False

def ageSelected(event): 
    e8.configure(state='normal')
    #e8.delete(0,"end")
    #e8.insert(0, int(ageCombo.current()))
    sv3.set(int(ageCombo.current()))
    e8.configure(state='disabled')
    
    for widget in myFrame2.winfo_children():
        widget.destroy()
    
    myFrame2.lowRange = []
    myFrame2.highRange = []
    myFrame2.totalList = []
    myFrame2.PercList = []
    vcmdAgeInt = (myFrame2.register(validateAge),'%P')

    CustTK.CTkLabel(myFrame2, text="Age Range").grid(row=0,column=2,columnspan=2)
    CustTK.CTkLabel(myFrame2, text="Total in range").grid(row=0,column=4)
    CustTK.CTkLabel(myFrame2, text="Age Percent").grid(row=0,column=5)  

    for x in range(int(ageCombo.get())):
        age_entry1 = CustTK.CTkEntry(myFrame2, width=75, border_width=2, corner_radius=5, validate="key", validatecommand=(vcmdAgeInt))
        age_entry2 = CustTK.CTkEntry(myFrame2, width=75, border_width=2, corner_radius=5, validate="key", validatecommand=(vcmdAgeInt))
        age_count = CustTK.CTkEntry(myFrame2) 
        age_perc = CustTK.CTkEntry(myFrame2)
        CustTK.CTkLabel(myFrame2, text="Age " + str(x+1)).grid(row=x+1,column=1)
        age_entry1.grid(row=x+1, column=2, pady=10, padx=5)
        age_entry2.grid(row=x+1, column=3, pady=10, padx=5)
        age_count.grid(row=x+1, column=4, pady=10, padx=5)
        age_count.configure(state='disabled')
        age_perc.grid(row=x+1, column=5, pady=10, padx=5)
        age_perc.configure(state='disabled')
        myFrame2.lowRange.append(age_entry1)
        myFrame2.highRange.append(age_entry2)
        myFrame2.totalList.append(age_count)
        myFrame2.PercList.append(age_perc)
        
    
    frame2Button = CustTK.CTkButton(myFrame2, text="Calculate Age Percentages/Save Settings", command=calcAge)
    frame2Button.grid(row=0,column=6, sticky=E)  
    

#AGE SETTINGS MAIN
AGEOPTIONS = [
    "0",
    "1",
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

ageCensus = {
    0: 392096,
    1: 392359,
    2: 392203,
    3: 392871,
    4: 391182,
    5: 392663,
    6: 397671,
    7: 400545,
    8: 401416,
    9: 392618,
    10: 382639,
    11: 376385,
    12: 378146,
    13: 373296,
    14: 375559,
    15: 383259,
    16: 398528,
    17: 405243,
    18: 426149,
    19: 453316,
    20: 477435,
    21: 489793,
    22: 493113,
    23: 500853,
    24: 507870,
    25: 513986,
    26: 518524,
    27: 502232,
    28: 489514,
    29: 492868,
    30: 503133,
    31: 507492,
    32: 506931,
    33: 507541,
    34: 505075,
    35: 509183,
    36: 501541,
    37: 488874,
    38: 479150,
    39: 477387,
    40: 476394,
    41: 475004,
    42: 462550,
    43: 462660,
    44: 468774,
    45: 489398,
    46: 485457,
    47: 481656,
    48: 477395,
    49: 481343,
    50: 507236,
    51: 541036,
    52: 556299,
    53: 559700,
    54: 547047,
    55: 552885,
    56: 542074,
    57: 528672,
    58: 521737,
    59: 507877,
    60: 491548,
    61: 484039,
    62: 462819,
    63: 439886,
    64: 421789,
    65: 411964,
    66: 401360,
    67: 390978,
    68: 387962,
    69: 383459,
    70: 327091,
    71: 300511,
    72: 288116,
    73: 272790,
    74: 250539,
    75: 236443,
    76: 218630,
    77: 206479,
    78: 193485,
    79: 179788,
    80: 171547,
    81: 159939,
    82: 148080,
    83: 141404,
    84: 132453,
    85: 123262,
    86: 110831,
    87: 97186,
    88: 86857,
    89: 74959,
    90: 65054,
    91: 54872,
    92: 45300,
    93: 36559,
    94: 29265,
    95: 22049,
    96: 14973,
    97: 9150,
    98: 5841,
    99: 3990
}

ageCombo = ttk.Combobox(myFrame1,value=AGEOPTIONS, justify=CENTER)
ageCombo.current(0)
ageCombo.bind("<<ComboboxSelected>>", ageSelected)
ageCombo.grid(row=3,column=6, sticky=W)


#myFrame2.entry_widgets = [myFrame2.create_entry_widget(x) for x in range(myFrame2.n)]

#def create_entry_widget(myFrame2, x):
#    new_widget = Entry(myFrame2.master)
#    new_widget.pack()
#    new_widget.insert(0, x)
#    return new_widget

#CREATING XLS
#def generateRegion(nSize,regInput):


def generateForm():

    #Calculate total number of variables in order to accurately place them in their respective rows.
    indice = 0
    
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
    # if (genVar == 1):
    #     indice = indice + genCount
    # if (ageVar == 1):
    #     indice = indice + ageCount 
    # if (regVar == 1):
    #     indice = indice + regCount
    genVar = int(var1.get())
    ageVar = int(var2.get())
    regVar = int(var3.get())
    interlockVar = int(var4.get())
    cellVar = int(cellBalType.get())
    cellCount = int(e_cb1.get())
    genCount = int(e7.get())
    ageCount = int(e8.get())
    regCount = int(e9.get())    
    cellQualed = int(var13.get())
    #Gender
    if (genVar == 1):
        worksheet1 = workbook.add_worksheet('Gender Quota')
        worksheet1.write('A1', '#=Gender Quota')
        indice = indice + genCount
        #print (indice)
        for count,eachRow in enumerate(range(indice)):
            worksheet.write('A' + str(eachRow+2), e1.get() + str(count+1))
            worksheet.write('B' + str(eachRow+2), e4.get() + ".r" + str(count+1))
            worksheet.write('C' + str(eachRow+2), myFrame4.entries[count].get())

            worksheet1.write('A' + str(count+2), e1.get() + str(count+1))
            worksheet1.write('B' + str(count+2), GenderQuotaCalc(count))
        print (myFrame4.genArray)
    #Age
    if (ageVar == 1):
        worksheet2 = workbook.add_worksheet('Age Quota')
        worksheet2.write('A1', '#=Age Quota')
        
        for count,eachRow in enumerate(range(indice+2,indice+2+ageCount)):          
            worksheet.write('A' + str(eachRow), e2.get() + str(count+1))        
            worksheet.write('B' + str(eachRow), e5.get() + ".r" + str(count+1))
            worksheet.write('C' + str(eachRow), myFrame2.lowRange[count].get() + "-" + myFrame2.highRange[count].get())

            worksheet2.write('A' + str(count+2), e2.get() + str(count+1))
            worksheet2.write('B' + str(count+2), AgeQuotaCalc(count))            
        indice = indice + ageCount
        print (myFrame2.AgePercList)
    #Region
    if (regVar == 1):
        worksheet3 = workbook.add_worksheet('Region Quota')
        worksheet3.write('A1', '#=Region Quota')

        for count, eachRow in enumerate(range(indice+2,indice+2+regCount)):
            worksheet.write('A' + str(eachRow), e3.get() + str(count+1))    
            worksheet.write('B' + str(eachRow), e6.get() + ".r" + str(count+1))
            if not(regCombo.current() == 10):
                worksheet.write('C' + str(eachRow), (regDict[regCombo.current()])[count]) 
            else:
                worksheet.write('C' + str(eachRow), myFrame3.entries[count].get()) 

            worksheet3.write('A' + str(count+2), e3.get() + str(count+1))    
            worksheet3.write('B' + str(count+2), RegQuotaCalc(count))    
        indice = indice + regCount
        print (myFrame3.regArray)
    
    #Cross Table cell balance MONADIC
    if (cellVar == 1):
        worksheet5 = workbook.add_worksheet('CELL BALANCE')
        worksheet5.write('A1', '#=CELL BALANCE')
        cbIndice = cellCount
        if cellQualed==1:
            worksheet5.write('B1', '#')
        rowLabels = e_cb4.get().split(",")
        
        #writes the defines for both cell/cellplus depending on if cells have qualifications
        for count, eachRow in enumerate(range(indice+2,indice+2+cellCount)):
            worksheet.write('A' + str(eachRow), "CELL" + str(count+1)) 

            if cellQualed == 1:
                worksheet.write('B' + str(eachRow), e_cb3.get() + "." + rowLabels[count])
                worksheet.write('A' + str(eachRow+cellCount), "CELLplus" + str(count+1))
                worksheet.write('B' + str(eachRow+cellCount), "plus")
                worksheet.write('C' + str(eachRow+cellCount), "CELL " + str(count+1)) 
            else:
                worksheet.write('B' + str(eachRow), "plus")
            worksheet.write('C' + str(eachRow), "CELL " + str(count+1)) 
        
        #writes the actual cell balance and all balancing factors       
        for count,eachRow in enumerate(range(1,cellCount+1)):
            worksheet5.write('A' + str(count+2), "CELL" + str(eachRow))
            if cellQualed == 1:
                worksheet5.write('B' + str(count+2), "CELLplus" + str(eachRow))    
                worksheet5.write('C' + str(count+2), "inf")
            else:
                worksheet5.write('B' + str(count+2), "inf")
        
        #writes cell balanced by region
        if (regBalance2.get() == 1):
            for count,eachRow in enumerate(range(1,cellCount+1)):
                worksheet5.write('B' + str(4+cbIndice+(regCount*count)), "CELL" + str(eachRow))
                worksheet5.write('C' + str(4+cbIndice+(regCount*count)), "CELLplus" + str(eachRow))
                for rCount, eachRow1 in enumerate(range(1,regCount+1)):
                    worksheet5.write('A' + str(rCount+4+cbIndice+(regCount*count)), e3.get() + str(eachRow1))   
                    if cellQualed == 1: 
                        worksheet5.write('D' + str(rCount+4+cbIndice+(regCount*count)), "inf")
                    else:
                        worksheet5.write('C' + str(rCount+4+cbIndice+(regCount*count)), "inf")
            worksheet5.write('A' + str(3+cbIndice), '#=Cell x Region')
            worksheet5.write('B' + str(3+cbIndice), '#')
            if var13.get() == 1:
                worksheet5.write('C' + str(3+cbIndice), '#')
            cbIndice = cbIndice + (regCount*cellCount) + (regBalance2.get()*2)

        #writes cell balanced by age
        if (ageBalance2.get() == 1):
            for count,eachRow in enumerate(range(1,cellCount+1)):
                worksheet5.write('B' + str(4+cbIndice+(ageCount*count)), "CELL" + str(eachRow))
                worksheet5.write('C' + str(4+cbIndice+(ageCount*count)), "CELLplus" + str(eachRow))
                for aCount, eachRow1 in enumerate(range(1,ageCount+1)):
                    worksheet5.write('A' + str(aCount+4+cbIndice+(ageCount*count)), e2.get() + str(eachRow1))   
                    if cellQualed == 1: 
                        worksheet5.write('D' + str(aCount+4+cbIndice+(ageCount*count)), "inf")
                    else:
                        worksheet5.write('C' + str(aCount+4+cbIndice+(ageCount*count)), "inf")
            worksheet5.write('A' + str(3+cbIndice), '#=Cell x Age')
            worksheet5.write('B' + str(3+cbIndice), '#')
            if var13.get() == 1:
                worksheet5.write('C' + str(3+cbIndice), '#')
            cbIndice = cbIndice + (ageCount*cellCount) + (ageBalance2.get()*2)


        #writes cell balanced by age
        if (genBalance2.get() == 1):
            for count,eachRow in enumerate(range(1,cellCount+1)):
                worksheet5.write('B' + str(4+cbIndice+(genCount*count)), "CELL" + str(eachRow))
                worksheet5.write('C' + str(4+cbIndice+(genCount*count)), "CELLplus" + str(eachRow))
                for gCount, eachRow1 in enumerate(range(1,genCount+1)):
                    worksheet5.write('A' + str(gCount+4+cbIndice+(genCount*count)), e1.get() + str(eachRow1))   
                    if cellQualed == 1: 
                        worksheet5.write('D' + str(gCount+4+cbIndice+(genCount*count)), "inf")
                    else:
                        worksheet5.write('C' + str(gCount+4+cbIndice+(genCount*count)), "inf")
            worksheet5.write('A' + str(3+cbIndice), '#=Cell x Age')
            worksheet5.write('B' + str(3+cbIndice), '#')
            if var13.get() == 1:
                worksheet5.write('C' + str(3+cbIndice), '#')
            cbIndice = cbIndice + (genCount*cellCount) + (genBalance2.get()*2)
            
        indice = indice + cellCount
        #print (myFrame3.regArray)

    #Legacy cell balance MONADIC
    if (cellVar == 2):
        worksheet5 = workbook.add_worksheet('CELL BALANCE')
        worksheet5.write('A1', '#=CELL BALANCE')
        for count, eachRow in enumerate(range(2,2+regCount*genCount*ageCount*cellCount,genCount*ageCount*cellCount)):
            worksheet5.write('A' + str(eachRow), e3.get() + str(count+1)) 
            for count2, eachRow2 in enumerate(range(2+count*genCount*ageCount*cellCount,2+(count+1)*genCount*ageCount*cellCount,genCount*cellCount)):
                worksheet5.write('B' + str(eachRow2), e2.get() + str(count2+1))             
                for count3, eachRow3 in enumerate(range(2+count*genCount*cellCount,2+(count+1)*genCount*cellCount,cellCount)):
                    worksheet5.write('C' + str(eachRow3), e1.get() + str(count3+1))
                    for count4, eachRow4 in enumerate(range(count3*cellCount,((count3+1)*cellCount))):
                        print (2+(count*ageCount*genCount*cellCount)+eachRow4)
                        worksheet5.write('D' + str(2+(count*ageCount*genCount*cellCount)+eachRow4), "CELL" + str(count4+1))

    #Interlock Quota
    if (interlockVar == 1 and (genVar+ageVar+regVar == 3)):
        worksheet4 = workbook.add_worksheet('Interlock Quota')
        worksheet4.write('A1', '#=Interlock Quota')    
        #print (genCount)
        worksheet4.write('B1', '#')
        worksheet4.write('C1', '#')
        #The following nested loops generates the interlocked region/age/gender markers as needed
        for count, eachRow in enumerate(range(2,2+regCount*genCount*ageCount,genCount*ageCount)):
            worksheet4.write('A' + str(eachRow), e3.get() + str(count+1)) 
            for count2, eachRow2 in enumerate(range(2+count*genCount*ageCount,2+(count+1)*genCount*ageCount,genCount)):
                worksheet4.write('B' + str(eachRow2), e2.get() + str(count2+1))              
                for count3, eachRow3 in enumerate(range(count2*genCount,((count2+1)*genCount))):
                    InterlockArray.append(myFrame3.RegPercList[count]*myFrame2.AgePercList[count2]*myFrame4.GenPercList[count3])
                    worksheet4.write('C' + str(2+(count*ageCount*genCount)+eachRow3), e1.get() + str(count3+1))
                    worksheet4.write('D' + str(2+(count*ageCount*genCount)+eachRow3), round(InterlockArray[(count*ageCount*genCount)+eachRow3]*int(e10.get()),None)) 
                    
        print (InterlockArray)
                                    
    elif (interlockVar == 1 and (genVar+ageVar+regVar == 2)):
        if genVar+ageVar == 2 and regVar == 0:
            doubleInterlock(workbook, ageCount, genCount,e2.get(),e1.get(),myFrame2.AgePercList,myFrame4.GenPercList)
        elif genVar+regVar == 2 and ageVar == 0:
            doubleInterlock(workbook, regCount, genCount,e3.get(),e1.get(),myFrame3.RegPercList,myFrame4.GenPercList)
        elif ageVar+ageVar == 2 and genVar == 0:
            doubleInterlock(workbook, regCount, ageCount,e3.get(),e2.get(),myFrame3.RegPercList,myFrame2.AgePercList)
    # Finally, close the Excel file
    # via the close() method.
    workbook.close()

#Function to handle the double interlocked demographics
def doubleInterlock(workbook, inter1, inter2,variable1,variable2,perclist1,perclist2):
    print("in doubleInt Function")
    InterlockArray = []
    worksheet4 = workbook.add_worksheet('Interlock Quota')
    worksheet4.write('A1', '#=Interlock Quota')       
    worksheet4.write('B1', '#')
    for count, eachRow in enumerate(range(2,2+inter1*inter2,inter2)):
        worksheet4.write('A' + str(eachRow), variable1 + str(count+1)) 
        for count2, eachRow2 in enumerate(range(2+count*inter2,2+(count+1)*inter2)):
            InterlockArray.append(perclist1[count]*perclist2[count2])
            worksheet4.write('B' + str(eachRow2), variable2 + str(count2+1))
            worksheet4.write('C' + str(eachRow2), round(InterlockArray[eachRow2-2]*int(e10.get()),None))

#Following 3 functions calculates the gender, age and region hard quotas/soft qutoas respectively, also will store the % of each in an array for interlock quota calulation
def GenderQuotaCalc(count):
    if int(var11.get()==1):
        return round(float(myFrame4.genArray[count].get())*float(e10.get()),None)
    else:
        return round(float(myFrame4.genArray[count].get())*float(e10.get()),None)*10

def AgeQuotaCalc(count):
    if int(var12.get()==1):
        return round(float(myFrame2.AgePercList[count])*float(e10.get()),None)
    else:
        return round(float(myFrame2.AgePercList[count])*float(e10.get()),None)*10

def RegQuotaCalc(count):
    if int(var13.get()==1):
        return round(float(myFrame3.regArray[count].get())*float(e10.get()),None)
    else:
        return round(float(myFrame3.regArray[count].get())*float(e10.get()),None)*10

# def calcInterlock(array1,array2,array3):
#     if array1
mySubmit = CustTK.CTkButton(myFrame1, text="Generate Form", command=generateForm)
mySubmit.grid(row=50, sticky=S, column=40)

window.mainloop()