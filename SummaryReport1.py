#!/usr/bin/python
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
# import openpyxl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import os
from datetime import datetime

# plt.rcParams.update({'font.size': 22})
system_ver = 0.02
root = tk.Tk()
root.title('Summary Report v{}'.format(system_ver))
root.iconbitmap('summaryreporticon.ico')
root.geometry("400x350")

top = tk.Toplevel() #Creates the toplevel window

entry1 = tk.Entry(top) #Username entry
entry2 = tk.Entry(top) #Password entry
button1 = tk.Button(top, text="Login", command=lambda:command1()) #Login button
button2 = tk.Button(top, text="Cancel", command=lambda:command2()) #Cancel button
# label1 = Label(root, text="This is your main window and you can input anything you want here")

def command1():
    if entry1.get() == "user" and entry2.get() == "password": #Checks whether username and password are correct
        root.deiconify() #Unhides the root window
        top.destroy() #Removes the toplevel window

def command2():
    top.destroy() #Removes the toplevel window
    root.destroy() #Removes the hidden root window
    sys.exit() #Ends the script


entry1.pack() #These pack the elements, this includes the items for the main window
entry2.pack()
button1.pack()
button2.pack()
# label1.pack()

savepdf=tk.IntVar()
savexlsx=tk.IntVar()
db_filename = "CARD USER I.D LIST 2.xlsx"
if (os.path.isfile(db_filename)):
    db = pd.read_excel(db_filename)
    print(db)
else:
    mylabel = tk.Label(root, text="cannot find file: CARD USER I.D LIST 2.xlsx",wraplength=400)
    mylabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)
    # while True:
    #     pass

df = pd.DataFrame()
dfproc = pd.DataFrame()

def computeHrs(timein, timeout):
    myin = timein.split(":")
    myout = timeout.split(":")

    hrs = int(myout[0])-int(myin[0])
    if hrs<0:
        return hrs+24
    else:
        return hrs

def billableHrs(timein, timeout):
    myin = timein.split(":")
    myout = timeout.split(":")
    inhr = int(myin[0])
    inmin = int(myin[1])
    outhr = int(myout[0])
    outmin = int(myout[1])
    hrs = outhr-inhr
    # print("hrs:{}".format(hrs))
    
    if inhr>=5 and inhr<8:
        shift = 1 #am shift
        if inhr==7 and inmin>15 and inmin<=59:
            late = 1
        else:
            late = 0
    elif inhr>=17 and inhr<20:
        shift = 2 #pm shift
        if inhr==19 and inmin>15 and inmin<=59:
            late = 1
        else:
            late = 0
    else:
        return 0
    
    if shift==1:
        if outhr>=19 and outhr<=23:
            return 12-late
        elif outhr>=13 and outhr<19:
            time=outhr-7-late #halfday atleaest 6 hrs 
            if time >=6:
                return 6
            else:
                return 0
        else:
            return 0
    elif shift==2:
        if outhr >=7 and outhr<=11:
            return 12-late
        elif outhr>=1 and outhr<7:
            time=outhr-19-late+24 #halfday atleaest 6 hrs 
            if time >=6:
                return 6
            else:
                return 0
        else:
            return 0

def openFile():
    global df
    global dfproc
    filename = fd.askopenfilename()
    
    hello = "file: " + filename
    myfilenamelabel.config(text=hello)
    statuslabel.config(text='status: File opened, ready to generate Report')
    excellabel.config(text='')
    pdflabel.config(text='')
    # myLabel.pack(pady=1)
    # wb = openpyxl.load_workbook(filename)
    # ws = wb.active
    df = pd.read_excel(filename)

    # data = list(ws.iter_rows(values_only=True))
    print(df)
    print("one  that mattersz")
    df.rename(columns=df.iloc[6])
    # df.columns = df.iloc[6]
    # df2 = df.drop(df.index[0:6])
    dfproc = df.iloc[7:]
    dfproc.columns = df.iloc[6]
    dfproc = dfproc.dropna(axis=1)
    dfproc = dfproc.reset_index()
    dfproc = dfproc.drop('index', axis=1)
    dfproc = dfproc.sort_values(by=["Name","Date"])

    
    # print(dfproc)
    # print("list of guards")
    # testguy = 11
    # print(dfproc['Name'])
    # print("si {}:".format(testguy))
    # print(dfproc['Name'][testguy])
    # employeehrs = computeHrs(dfproc['Time In'][testguy],dfproc['Time Out'][testguy])
    # billablehrs = billableHrs(dfproc['Time In'][testguy],dfproc['Time Out'][testguy])
    # print("tothrs:{} billable:{}".format(employeehrs,billablehrs))

    
def generateReport():
    global dfproc,savexlsx,savepdf
    print("generating...")
    dfproc['Total Hours'] = " "
    dfproc['Billable Hours'] = " "
    for index, row in dfproc.iterrows():
        myhrs = computeHrs(row['Time In'],row['Time Out'])
        billablehrs = billableHrs(row['Time In'],row['Time Out'])
        row['Total Hours']=myhrs
        row['Billable Hours']=billablehrs
    
    dfproc = dfproc.reset_index()
    dfproc = dfproc.drop('index', axis=1) 
    print(dfproc)
    col = dfproc.pop("Name")
    dfproc.insert(0,col.name,col)
    print("moving aroundd")
    dfproc['Gender'] = " "
    dfproc['Remarks'] = " "
    dfproc['Day'] = " "
    for index, row in dfproc.iterrows():
        df_new = db[db['Card N°.'] == row['Card No']]
        df_new = df_new.reset_index()
        if df_new.empty:
            row['Gender'] = " "
        else:
            row['Gender']=df_new['Gender'][0]
        myday = datetime.strptime(row['Date'],'%Y-%m-%d').strftime('%a')
        row['Day']=myday
    col = dfproc.pop("Card No")
    dfproc.insert(1,"Employee Number",col)
    col = dfproc.pop("Gender")
    dfproc.insert(2,col.name,col)
    col = dfproc.pop("Department")
    dfproc.insert(3,"Category",col)
    col = dfproc.pop("Day")
    dfproc.insert(5,col.name,col)
    print(dfproc)
    if savepdf.get()==1:
        pdflabel.config(text='pdf: generating...')
        savePdf(dfproc)
    if savexlsx.get()==1:
        excellabel.config(text='spreadsheet: generating...')
        saveXlsx(dfproc)
    
    dfproc.iloc[0:0]
    statuslabel.config(text='status: Report generated, open new file')
    myfilenamelabel.config(text='file: ')

def saveXlsx(dfproc):
    filename_out=fd.asksaveasfile(mode='w',defaultextension=".xlsx",filetypes=(("Excel Workbook", "*.xlsx"),("All Files", "*.*")))
    print(filename_out.name)
    dfproc.to_excel(filename_out.name,index=False)
    excellabel.config(text='spreadsheet: '+filename_out.name)

def set_align_for_column(table, col, align="left"):
    cells = [key for key in table._cells if key[1] == col]
    for cell in cells:
        table._cells[cell]._loc = align

def savePdf(dfproc):
    filename_out=fd.asksaveasfile(mode='w',defaultextension=".pdf",filetypes=(("PDF file", "*.pdf"),("All Files", "*.*")))
    print(filename_out.name)
    fig, ax =plt.subplots(figsize=(8.5,11))
    ax.axis('tight')
    ax.axis('off')
    the_table = ax.table(cellText=dfproc.values,colLabels=dfproc.columns,loc='center')
    the_table.auto_set_font_size(False)
    the_table.set_fontsize(6)
    the_table.auto_set_column_width(col=[0,1,2,3,5])
    # the_table[0, col]._text.set_horizontalalignment('left') 
    pp = PdfPages(filename_out.name)
    pp.savefig(fig, bbox_inches='tight')
    pp.close()
    pdflabel.config(text='pdf: '+filename_out.name)
    # os.startfile(filename_out.name)
# def saveFile():
    # print("saving...")
    # html_string = df.to_html()
    # pdfkit.from_string(html_string,"output_file.pdf")
    # print("file saved")
def print_selection():
    print("as pdf:{} as xlsx:{}".format(savepdf,savexlsx))

# e = tk.Entry(root, width=50, font=('Helvetica',20))
# e.pack(padx=10, pady=10)

# myButton = tk.Button(root, text="Enter naym",command=myClick)
# myButton.pack(padx=20)
myButton2 = tk.Button(root, text="Open File",command=openFile)
myButton2.pack(side=tk.TOP, anchor=tk.NW,padx=20,pady=10)
myfilenamelabel = tk.Label(root, text="file:",wraplength=400)
myfilenamelabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)

c1 = tk.Checkbutton(root, text='save as PDF',variable=savepdf, onvalue=1, offvalue=0, command=print_selection)
c1.pack(side=tk.TOP,anchor=tk.W,padx=20)
c2 = tk.Checkbutton(root, text='save as spreadsheet',variable=savexlsx, onvalue=1, offvalue=0, command=print_selection)
c2.pack(side=tk.TOP, anchor=tk.W,padx=20)

myButton = tk.Button(root, text="Generate Report",command=generateReport)
myButton.pack(side=tk.TOP, anchor=tk.W,padx=20,pady=10)
statuslabel = tk.Label(root, text="status:",wraplength=400, justify="left")
statuslabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)
pdflabel = tk.Label(root, text="",wraplength=400, justify="left")
pdflabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)
excellabel = tk.Label(root, text="",wraplength=400, justify="left")
excellabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)

root.withdraw()
root.mainloop()
