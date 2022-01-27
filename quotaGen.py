import enum
import pandas as pd
import xlsxwriter
from tkinter import *
from tkinter import ttk

window=Tk()
#add title for the window
window.title("Canview Quota Generator V1")
window.geometry("1000x800+150+250")

my_notebook = ttk.Notebook(window)
my_notebook.grid(row=0)

myFrame1 = Frame(my_notebook, width=900, height=800)
myFrame2 = Frame(my_notebook, width=900, height=800)
myFrame3 = Frame(my_notebook, width=900, height=800)
myFrame4 = Frame(my_notebook, width=900, height=800)
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
Label(myFrame1, text="Hard Quota:").grid(row=1, sticky=W)
Label(myFrame1, text="Demos Used:").grid(row=1, column=1, sticky=W)
Label(myFrame1, text="Number of variable groups").grid(row=1, column=7, sticky=E)
Label(myFrame1, text="TOTAL N: ").grid(row=1, column=4)
is_on=True
#defined function to enable 2nd set of interlocks when 1st interlock button clicked
def enableInt():
    global is_on

    if is_on:
        genButton2.config(state=NORMAL)
        ageButton2.config(state=NORMAL)
        regButton2.config(state=NORMAL)
        intButton2.config(state=NORMAL)
        is_on=False
    else:
        genButton2.deselect()
        ageButton2.deselect()
        regButton2.deselect()
        intButton2.deselect()

        genButton2.config(state=DISABLED)
        ageButton2.config(state=DISABLED)
        regButton2.config(state=DISABLED)
        intButton2.config(state=DISABLED)

        is_on=True       
#Main page - 1st set of interlock
var1 = IntVar()
Checkbutton(myFrame1, text="Gender", variable=var1).grid(row=2, column=1, sticky=W)
var2 = IntVar()
Checkbutton(myFrame1, text="Age", variable=var2).grid(row=3, column=1, sticky=W)
var3 = IntVar()
Checkbutton(myFrame1, text="Region", variable=var3).grid(row=4, column=1, sticky=W)
var4 = IntVar()
Checkbutton(myFrame1, text="Interlock 1", variable=var4, command=enableInt).grid(row=5, column=1, sticky=W)
#Main page - HQ checkboxes for 1st set of interlock
var11 = IntVar()
Checkbutton(myFrame1, variable=var11).grid(row=2)
var12 = IntVar()
Checkbutton(myFrame1, variable=var12).grid(row=3)
var13 = IntVar()
Checkbutton(myFrame1, variable=var13).grid(row=4)
var14 = IntVar()
Checkbutton(myFrame1, variable=var14).grid(row=5)
#Main page - 2nd set of interlock
var5 = IntVar()
genButton2 = Checkbutton(myFrame1, state=DISABLED, text="Gender", variable=var5)
genButton2.grid(row=2, column=8, sticky=W)
var6 = IntVar()
ageButton2 = Checkbutton(myFrame1, state=DISABLED, text="Age", variable=var6)
ageButton2.grid(row=3, column=8, sticky=W)
var7 = IntVar()
regButton2 = Checkbutton(myFrame1, state=DISABLED, text="Region", variable=var7)
regButton2.grid(row=4, column=8, sticky=W)
var8 = IntVar()
intButton2 = Checkbutton(myFrame1, state=DISABLED, text="Interlock 2", variable=var8)
intButton2.grid(row=5, column=8, sticky=W)


ttk.Separator(myFrame1, orient='horizontal').grid(row=6, columnspan=200, sticky="ew")


Label(myFrame1, text="Gender Marker: ").grid(row=2,column=2)
Label(myFrame1, text="Age Marker: ").grid(row=3,column=2)
Label(myFrame1, text="Region Marker: ").grid(row=4,column=2)

e1 = Entry(myFrame1, width=10, borderwidth=3)
e2 = Entry(myFrame1, width=10, borderwidth=3)
e3 = Entry(myFrame1, width=10, borderwidth=3)

e1.grid(row=2,column=3)
e2.grid(row=3,column=3)
e3.grid(row=4,column=3)

e1.insert(0, "gen")
e2.insert(0, "age")
e3.insert(0, "reg")


Label(myFrame1, text="Gender Var: ").grid(row=2,column=4)
Label(myFrame1, text="Age Var: ").grid(row=3,column=4)
Label(myFrame1, text="Region Var: ").grid(row=4,column=4)

e4 = Entry(myFrame1, width=25, borderwidth=3)
e5 = Entry(myFrame1, width=25, borderwidth=3)
e6 = Entry(myFrame1, width=25, borderwidth=3)
e10 = Entry(myFrame1, width=25, borderwidth=3)

e10.grid(row=1,column=5)
e4.grid(row=2,column=5)
e5.grid(row=3,column=5)
e6.grid(row=4,column=5)

e4.insert(0, "GENDER")
e5.insert(0, "QUOTAGERANGE")
e6.insert(0, "REGION")

sv1 = StringVar()
sv2 = StringVar()

e7 = Entry(myFrame1, width=3, borderwidth=3, textvariable=sv2)
e8 = Entry(myFrame1, width=3, borderwidth=3)
e9 = Entry(myFrame1, width=3, borderwidth=3, textvariable=sv1)

e7.grid(row=2,column=7)
e8.grid(row=3,column=7)
e9.grid(row=4,column=7)

e7.insert(0, 0)
e8.insert(0, 0)
e9.insert(0, 0)

e7.configure(state='disabled')
e8.configure(state='disabled')
e9.configure(state='disabled')
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

def callback2(sv):
    #if (genCombo.current() == 2):   
        myFrame4.entries=[]

        for widget in myFrame4.winfo_children():
            widget.destroy()

        for y in range(int(sv.get())):
            gen_entry = Entry(myFrame4)
            gen_name = Entry(myFrame4)
            
            #Create gender name text boxes
            gen_name.grid(row=y, column=2, pady=10, padx=5)
            #Fill gender name text boxes with default gender labels
            gen_name.insert(0, (genNames[y+1])) 
            #Create gender % text boxes
            gen_entry.grid(row=y, column=3, pady=10, padx=5)
            #Fill gender % text boxes with default gender %
            gen_entry.insert(0, (str(genPerc[y+1]*100)) + "%") 
            Label(myFrame4, text="Gen" + str(y+1)).grid(row=y,column=1)
            myFrame4.entries.append(gen_name) 

sv2.trace("w", lambda name, index, mode, sv=sv2: callback2(sv))

def genSelected(event):
    e7.configure(state='normal')
    e7.delete(0,"end")

    myFrame4.labels=[]

    for widget in myFrame4.winfo_children():
        widget.destroy()

    e7.insert(0,(int(genDict[genCombo.current()])))
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

genCombo = ttk.Combobox(myFrame1,value=GENOPTIONS)
genCombo.current(0)
genCombo.bind("<<ComboboxSelected>>", genSelected)
genCombo.grid(row=2,column=6, sticky=W)

#REG TAB
regDict = {
  1: ["WEST","ONTARIO","QC","ATLANTIC"],
  2: ["WEST","ONTARIO","EN QC","ATLANTIC"],
  3: ["WEST","ONTARIO","FR QC","ATLANTIC"],
  4: ["WEST","ONTARIO","ATLANTIC"],
  5: ["NORTHEAST","MIDWEST","SOUTH","WEST"],
  6: [""],
  7: ["BC","PRAIRIES","ONTARIO","EN QC","FR QC","ATLANTIC"],
  8: ["BC","PRAIRIES","ONTARIO","EN QC","ATLANTIC"],
  9: ["BC","PRAIRIES","ONTARIO","FR QC","ATLANTIC"],
  10: ["BC","PRAIRIES","ONTARIO","ATLANTIC"]
}
regPerc = {
  1: [0.3012,0.3898,0.2362,0.0728],
  2: [0.3850,0.4980,0.0240,0.0930],
  3: [0.3106,0.4020,0.2122,0.0752],
  4: [0.3942,0.5104,0.0954],
  5: [0.178186900785363,0.215534586540639,0.372431791932523,0.233846720741475],
  6: [0],
  7: [0.1323,0.1688,0.3898,0.024,0.2122,0.0729],
  8: [0.169,0.216,0.498,0.024,0.093],
  9: [0.1365,0.1741,0.402,0.2122,0.0752],
  10: [0.1732,0.221,0.5104,0.0954]
}
def callback1(sv):
    if (regCombo.current() == 6):
        myFrame3.entries=[]
    
        for widget in myFrame3.winfo_children():
            widget.destroy()
        Label(myFrame3, text="Region Label").grid(row=0,column=1)
        Label(myFrame3, text="Region Percent").grid(row=0,column=2)    

        for y in range(int(sv.get())):
            reg_entry = Entry(myFrame3)
            reg_name = Entry(myFrame3)
            
            reg_name.grid(row=y+1, column=1, pady=10, padx=5)
            reg_entry.grid(row=y+1, column=2, pady=10, padx=5)
            reg_name.insert(0, "REGION " + str(y+1))
            myFrame3.entries.append(reg_name) 

sv1.trace("w", lambda name, index, mode, sv=sv1: callback1(sv))
def checkBCBreak():
    print (regCombo.current())

def regSelected(event):
    e9.configure(state='normal')
    e9.delete(0,"end")

    for widget in myFrame3.winfo_children():
        widget.destroy()

    Label(myFrame3, text="Region Label").grid(row=0,column=1)
    Label(myFrame3, text="Region Percent").grid(row=0,column=2)
    #BC Breaks variable set
    rvar1 = IntVar()
    Checkbutton(myFrame3, text="BC breaks", variable=rvar1, command=checkBCBreak).grid(row=0, column=10, sticky=W)       
    #print (rvar1.get())

    if not(regCombo.current() == 6):
        startIndex = regCombo.current()
        if rvar1.get() == 0:
            startIndex = regCombo.current()
        else:
            startIndex = regCombo.current()+6

        e9.insert(0, len(regDict[startIndex]))
        e9.configure(state='disabled')
        
        for x in range(len(regDict[startIndex])):
            reg_entry = Entry(myFrame3)
            Label(myFrame3, text=str((regDict[startIndex])[x])).grid(row=x+1,column=1)
            reg_entry.grid(row=x+1, column=2, pady=10, padx=5)
            reg_entry.insert(0, ((regPerc[startIndex])[x])) 
    else:
        e9.insert(0, 6)


#Refresh button to check if 'bc breaks' checked off
#myRefresh = Button(myFrame3, text="REFRESH", command=regSelected)
#myRefresh.grid(sticky=E)

#REGION SETTINGS
REGOPTIONS = [
    "SELECT ONE",
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
        myFrame2.percList[j].configure(state='normal')
        myFrame2.percList[j].insert(0, str((int(myFrame2.totalList[j].get())/tempTotal)*100) + "%")
        myFrame2.percList[j].configure(state='disabled')

def ageSelected(event): 
    e8.configure(state='normal')
    e8.delete(0,"end")
    e8.insert(0, int(ageCombo.current()))
    e8.configure(state='disabled')
    
    for widget in myFrame2.winfo_children():
        widget.destroy()
    
    myFrame2.lowRange = []
    myFrame2.highRange = []
    myFrame2.totalList = []
    myFrame2.percList = []
    Label(myFrame2, text="Age Range").grid(row=0,column=2,columnspan=2)
    Label(myFrame2, text="Total in range").grid(row=0,column=4)
    Label(myFrame2, text="Age Percent").grid(row=0,column=5)  

    for x in range(int(ageCombo.get())):
        age_entry1 = Entry(myFrame2,width=5)
        age_entry2 = Entry(myFrame2,width=5)
        age_count = Entry(myFrame2) 
        age_perc = Entry(myFrame2)
        Label(myFrame2, text="Age " + str(x+1)).grid(row=x+1,column=1)
        age_entry1.grid(row=x+1, column=2, pady=10, padx=5)
        age_entry2.grid(row=x+1, column=3, pady=10, padx=5)
        age_count.grid(row=x+1, column=4, pady=10, padx=5)
        age_count.configure(state='disabled')
        age_perc.grid(row=x+1, column=5, pady=10, padx=5)
        age_perc.configure(state='disabled')
        myFrame2.lowRange.append(age_entry1)
        myFrame2.highRange.append(age_entry2)
        myFrame2.totalList.append(age_count)
        myFrame2.percList.append(age_perc)
    
    frame2Button = Button(myFrame2, text="Calculate Age Percentages", command=calcAge)
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

ageCombo = ttk.Combobox(myFrame1,value=AGEOPTIONS)
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
    genCount = int(e7.get())
    ageCount = int(e8.get())
    regCount = int(e9.get())    
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
            worksheet1.write('B' + str(count+2), "inf")

    #Age
    if (ageVar == 1):
        worksheet2 = workbook.add_worksheet('Age Quota')
        worksheet2.write('A1', '#=Age Quota')
        
        for count,eachRow in enumerate(range(indice+2,indice+2+ageCount)):          
            worksheet.write('A' + str(eachRow), e2.get() + str(count+1))        
            worksheet.write('B' + str(eachRow), e5.get() + ".r" + str(count+1))
            worksheet.write('C' + str(eachRow), myFrame2.lowRange[count].get() + "-" + myFrame2.highRange[count].get())

            worksheet2.write('A' + str(count+2), e2.get() + str(count+1))
            worksheet2.write('B' + str(count+2), "inf")            
        indice = indice + ageCount

    #Region
    if (regVar == 1):
        worksheet3 = workbook.add_worksheet('Region Quota')
        worksheet3.write('A1', '#=Region Quota')

        for count, eachRow in enumerate(range(indice+2,indice+2+regCount)):
            worksheet.write('A' + str(eachRow), e3.get() + str(count+1))    
            worksheet.write('B' + str(eachRow), e6.get() + ".r" + str(count+1))
            if not(regCombo.current() == 6):
                worksheet.write('C' + str(eachRow), (regDict[regCombo.current()])[count]) 
            else:
                worksheet.write('C' + str(eachRow), myFrame3.entries[count].get()) 

            worksheet3.write('A' + str(count+2), e3.get() + str(count+1))    
            worksheet3.write('B' + str(count+2), "inf")    
        indice = indice + regCount

    if (interlockVar == 1 and (genVar+ageVar+regVar == 3)):
        worksheet4 = workbook.add_worksheet('Interlock Quota')
        worksheet4.write('A1', '#=Interlock Quota')    
        #print (genCount)
        worksheet4.write('B1', '#')
        worksheet4.write('C1', '#')
        for count, eachRow in enumerate(range(2,2+regCount*genCount*ageCount,genCount*ageCount)):
            worksheet4.write('A' + str(eachRow), e3.get() + str(count+1)) 
            for count2, eachRow2 in enumerate(range(2+count*genCount*ageCount,2+(count+1)*genCount*ageCount,genCount)):
                worksheet4.write('B' + str(eachRow2), e2.get() + str(count2+1))
                
                for count3, eachRow3 in enumerate(range(2+count2*genCount,(2+(count2+1)*genCount))):
                    print ("EXCEL INDEX: " + str(eachRow+eachRow2+eachRow3))
                    print ("MARKER: " + e1.get() + str(count3+1))
                    print ("COUNT1: " + str(count))
                    print ("COUNT2: " + str(count2))
                    print (range(2+count2*genCount,(2+(count2+1)*genCount)))
                    print ("-------------------------------")
                    worksheet4.write('C' + str(eachRow+eachRow2+eachRow3), e1.get() + str(count3+1))
                    worksheet4.write('D' + str(eachRow3), "inf")                 
        # myCountStart = 2
        # myCountEnd = genCount*ageCount*regCount
        
        # while myCountStart <= myCountEnd:
        #     for eachRow in range(0,genCount):
        #         worksheet4.write('C' + str(myCountStart), e1.get() + str(eachRow+1))
        #         worksheet4.write('D' + str(myCountStart), "inf")  
        #         myCountStart+=1

 
    elif (interlockVar == 1 and (genVar+ageVar+regVar == 2)):
        if genVar+ageVar == 2 and regVar == 0:
            doubleInterlock(workbook, ageCount, genCount,e2.get(),e1.get())
        elif genVar+regVar == 2 and ageVar == 0:
            doubleInterlock(workbook, regCount, genCount,e3.get(),e1.get())
        elif ageVar+ageVar == 2 and genVar == 0:
            doubleInterlock(workbook, regCount, ageCount,e3.get(),e2.get())
    # Finally, close the Excel file
    # via the close() method.
    workbook.close()

def doubleInterlock(workbook, inter1, inter2,variable1,variable2):
    print("in doubleInt Function")
    worksheet4 = workbook.add_worksheet('Interlock Quota')
    worksheet4.write('A1', '#=Interlock Quota')       
    worksheet4.write('B1', '#')
    for count, eachRow in enumerate(range(2,2+inter1*inter2,inter2)):
        worksheet4.write('A' + str(eachRow), variable1 + str(count+1)) 
        for count2, eachRow2 in enumerate(range(2+count*inter2,2+(count+1)*inter2)):
            worksheet4.write('B' + str(eachRow2), variable2 + str(count2+1))

mySubmit = Button(myFrame1, text="Generate Form", command=generateForm)
mySubmit.grid(row=50, sticky=S, column=2)

window.mainloop()