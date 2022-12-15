

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
root = tk.Tk()
root.title('Summary Report')
root.iconbitmap('summaryreporticon.ico')
root.geometry("400x500")


top = tk.Toplevel() #Creates the toplevel window
top.geometry("400x500")
headerlabel = tk.Label(top, text="Kenichi Security (Okada)",wraplength=400, justify="left",font=('Helvetica',15)).pack(side=tk.TOP,anchor=tk.NW,padx=20,pady=20)
tk.ttk.Separator(top, orient='horizontal').pack(fill='x',pady=5)
tk.Label(top, text="User Login",wraplength=400, justify="left",font=('Helvetica',13)).pack(side=tk.TOP,anchor=tk.NW,padx=20,pady=20)

entry1 = tk.Entry(top) #Username entry
entry2 = tk.Entry(top,show="*") #Password entry
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
button1.pack(pady=10)
button2.pack()
# label1.pack()

savepdf=tk.IntVar(value=1)
savexlsx=tk.IntVar(value=1)
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
def isNaN(string):
    return string != string

def computeHrs(timein, timeout):
    if isNaN(timein) or isNaN(timeout):
        return 0
    myin = timein.split(":")
    myout = timeout.split(":")

    hrs = int(myout[0])-int(myin[0])
    if hrs<0:
        return hrs+24
    else:
        return hrs

def billableHrs(timein, timeout):
    if isNaN(timein) or isNaN(timeout):
        return 0
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
    result = df.head(10)
    print("First 10 rows of the DataFrame:")
    print(result)
    
    dfproc = df.iloc[7:]
    dfproc.columns = df.iloc[6]
    print(dfproc.columns)
    dfproc = dfproc.loc[:,dfproc.columns.notna()]
    # dfproc = dfproc.dropna(axis=1)
    dfproc = dfproc.reset_index()
    dfproc = dfproc.drop('index', axis=1)
    dfproc = dfproc.sort_values(by=["Name","Date"])
    result = dfproc.head(10)
    print("First 10 rows of the DataFrame:")
    print(result)
    
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
    selection = radbut.get()
    print(selection)
    if selection ==0:
        generateDailyReport()
    if selection ==1:
        generateDailyReport()
    if selection ==2:
        generateCEZAReport()

def generateCEZAReport():
    global dfproc,savexlsx,savepdf
    # dfproc=db
    dfproc.iloc[0:0]
    dfproc = db[['USER ID', 'Name', 'Gender', 'Staff No', 'Department']].copy()
    dfproc = dfproc.drop([0])
    dfproc = dfproc.reset_index()
    dfproc = dfproc.drop('index', axis=1) 
    print(dfproc)
    statuslabel.config(text='status: CEZA Report generated')
    myfilenamelabel.config(text='file: ')
    if savepdf.get()==1:
        pdflabel.config(text='pdf: generating...')
        savePdf(dfproc,2)
    if savexlsx.get()==1:
        excellabel.config(text='spreadsheet: generating...')
        saveXlsx(dfproc,2)
    
    dfproc.iloc[0:0]
    df.iloc[0:0]
    
    if savepdf.get()==1:
        statuslabel.config(text='status: opening pdf...')
        outfilename = pdflabel.cget("text")
        os.startfile(outfilename[5:])
    if savexlsx.get()==1:
        statuslabel.config(text='status: opening spreadheet...')
        outfilename = excellabel.cget("text")
        os.startfile(outfilename[13:])

    statuslabel.config(text='status: Done. Open new file')
    myfilenamelabel.config(text='file: ')

def generateDailyReport():
    global dfproc,savexlsx,savepdf
    print(dfproc)
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
    dfproc = dfproc.fillna('')
    print(dfproc)
    statuslabel.config(text='status: Reports generated')
    myfilenamelabel.config(text='file: ')
    if savepdf.get()==1:
        pdflabel.config(text='pdf: generating...')
        savePdf(dfproc,0)
    if savexlsx.get()==1:
        excellabel.config(text='spreadsheet: generating...')
        saveXlsx(dfproc,0)
    
    dfproc.iloc[0:0]
    df.iloc[0:0]
    
    if savepdf.get()==1:
        statuslabel.config(text='status: opening pdf...')
        outfilename = pdflabel.cget("text")
        os.startfile(outfilename[5:])
    if savexlsx.get()==1:
        statuslabel.config(text='status: opening spreadheet...')
        outfilename = excellabel.cget("text")
        os.startfile(outfilename[13:])

    statuslabel.config(text='status: Done. Open new file')
    myfilenamelabel.config(text='file: ')

def saveXlsx(dfproc,reporttype):
    now = datetime.now()
    if reporttype==0:
        initfile = "SummaryDailyReport_"+now.strftime("%d-%m-%Y")
    elif reporttype==1:
        initfile = "SummaryBiMonthlyReport_"+now.strftime("%d-%m-%Y")
    elif reporttype==2:
        initfile = "SummaryCEZAReport_"+now.strftime("%d-%m-%Y")
    filename_out=fd.asksaveasfile(mode='w',defaultextension=".xlsx",initialfile=initfile,filetypes=(("Excel Workbook", "*.xlsx"),("All Files", "*.*")))
    print(filename_out.name)
    writer = pd.ExcelWriter(filename_out.name, engine='xlsxwriter')
    dfproc.to_excel(writer,index=False, sheet_name='Sheet1')
    for column in dfproc:
        column_width = max(dfproc[column].astype(str).map(len).max(), len(column))
        col_idx = dfproc.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

    writer.save()

    excellabel.config(text='spreadsheet: '+filename_out.name)

def savePdf(dfproc,reporttype):

    now = datetime.now()
    if reporttype==0:
        initfile = "SummaryDailyReport_"+now.strftime("%d-%m-%Y")
    elif reporttype==1:
        initfile = "SummaryBiMonthlyReport_"+now.strftime("%d-%m-%Y")
    elif reporttype==2:
        initfile = "SummaryCEZAReport_"+now.strftime("%d-%m-%Y")
    filename_out=fd.asksaveasfile(mode='w',defaultextension=".pdf",initialfile=initfile,filetypes=(("PDF file", "*.pdf"),("All Files", "*.*")))
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
headerlabel = tk.Label(root, text="Kenichi Security (Okada)",wraplength=400, justify="left",font=('Helvetica',15)).pack(side=tk.TOP,anchor=tk.NW,padx=20,pady=20)
tk.ttk.Separator(root, orient='horizontal').pack(fill='x',pady=5)
myButton2 = tk.Button(root, text="Open File",command=openFile,font=('Helvetica',13))
myButton2.pack(side=tk.TOP, anchor=tk.NW,padx=20,pady=10)
myfilenamelabel = tk.Label(root, text="file:",wraplength=400)
myfilenamelabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)
tk.ttk.Separator(root, orient='horizontal').pack(fill='x',pady=5)

def ShowChoice():
    print(radbut.get())

radbut = tk.IntVar()
radbut.set(0)
radbut_values = [("Daily T&A Report", 0),
   	            ("Bi monthly Billing Report", 1),
                ("CEZA Report", 2)]
for text, value in radbut_values:
    tk.Radiobutton(root, 
                   text=text,
                   padx = 20, 
                   variable=radbut, 
                   command=ShowChoice,
                   value=value,font=('Helvetica',13)).pack(anchor=tk.W)
 

myButton = tk.Button(root, text="Generate Report",command=generateReport,font=('Helvetica',13))
myButton.pack(side=tk.TOP, anchor=tk.W,padx=20,pady=10)
c1 = tk.Checkbutton(root, text='save as PDF',variable=savepdf, onvalue=1, offvalue=0, command=print_selection,font=('Helvetica',13))
c1.pack(side=tk.TOP,anchor=tk.W,padx=20)
c2 = tk.Checkbutton(root, text='save as spreadsheet',variable=savexlsx, onvalue=1, offvalue=0, command=print_selection,font=('Helvetica',13))
c2.pack(side=tk.TOP, anchor=tk.W,padx=20)
tk.ttk.Separator(root, orient='horizontal').pack(fill='x',pady=5)

statuslabel = tk.Label(root, text="status:",wraplength=400, justify="left")
statuslabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)
pdflabel = tk.Label(root, text="",wraplength=400, justify="left")
pdflabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)
excellabel = tk.Label(root, text="",wraplength=400, justify="left")
excellabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)

root.withdraw()
root.mainloop()