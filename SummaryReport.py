import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
# import openpyxl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import os,sys
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings("ignore", category=FutureWarning) 

# plt.rcParams.update({'font.size': 22})
sys_ver = "0.13"
root = tk.Tk()
root.title('Summary Report v'+str(sys_ver))
root.iconbitmap('summaryreporticon.ico')
root.geometry("400x600")

def destroyer():
    # return 1
    root.quit()
    root.destroy()
    sys.exit()
def top_destroyer():
    # top.quit()
    # top.destroy()
    sys.exit()

top = tk.Toplevel() #Creates the toplevel window
top.geometry("400x600")
# top.protocol("WM_DELETE_WINDOW", top_destroyer)
headerlabel = tk.Label(top, text="Kenichi Security (Okada)",wraplength=400, justify="left",font=('Helvetica',15)).pack(side=tk.TOP,anchor=tk.NW,padx=20,pady=20)
tk.ttk.Separator(top, orient='horizontal').pack(fill='x',pady=5)
tk.Label(top, text="User Login",wraplength=400, justify="left",font=('Helvetica',13)).pack(side=tk.TOP,anchor=tk.NW,padx=20,pady=20)

def command1():
    if entry1.get() == "junjun" and entry2.get() == "avelino12345": #Checks whether username and password are correct
        root.deiconify() #Unhides the root window
        top.destroy() #Removes the toplevel window

def command2():
    top.destroy() #Removes the toplevel window
    root.destroy() #Removes the hidden root window
    sys.exit() #Ends the script

entry1 = tk.Entry(top) #Username entry
entry2 = tk.Entry(top,show="*") #Password entry
button1 = tk.Button(top, text="Login", command=lambda:command1()) #Login button
button2 = tk.Button(top, text="Cancel", command=lambda:command2()) #Cancel button
# label1 = Label(root, text="This is your main window and you can input anything you want here")

entry1.pack() #These pack the elements, this includes the items for the main window
entry2.pack()
button1.pack(pady=10)
button2.pack()
# label1.pack()
dropdown_month = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
month_30 = ["Apr","Jun","Sep","Nov",]
month_31 = ["Jan","Mar","May","Jul","Aug","Oct","Dec"]
dropdown_cutoff = ["1-15","16-31"] 
monthvariable = tk.StringVar(root)
monthvariable.set(dropdown_month[int(datetime.now().strftime('%m'))-1]) # default value
cutoffvariable = tk.StringVar(root)
if int(datetime.now().strftime('%d')) >15:
    cutoffvariable.set(dropdown_cutoff[0]) # default value
else:
    cutoffvariable.set(dropdown_cutoff[1]) # default value
# year_entry = tk.Entry(root)
yearvariable = tk.StringVar(root)
thisyr = datetime.now().strftime("%Y")
yearvariable.set(thisyr)
year_entry = tk.Entry(root,textvariable = yearvariable)
savepdf=tk.IntVar(value=1)
savexlsx=tk.IntVar(value=1)
bimoformat=tk.IntVar(value=1)
tinputformat=tk.IntVar(value=0)

db_filename = "CARD USER I.D LIST 2.xlsx"
if (os.path.isfile(db_filename)):
    db = pd.read_excel(db_filename,sheet_name='Sheet')
    db_sched = pd.read_excel(db_filename,sheet_name='Schedules')
    print(db)
    print(db_sched)
else:
    mylabel = tk.Label(root, text="cannot find file: CARD USER I.D LIST 2.xlsx",wraplength=400)
    mylabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)
    # while True:
    #     pass

df = pd.DataFrame()
dfproc = pd.DataFrame()
def isNaN(string):
    return string != string

def destroyer():
    root.quit()
    root.destroy()
    top.quit()
    top.destroy()
    sys.exit()

def computeHrs(timein, timeout):
    if isNaN(timein) or isNaN(timeout):
        return 0
    myin = datetime.strptime(timein,"%H:%M:%S")
    myout = datetime.strptime(timeout,"%H:%M:%S")
    delta = myout-myin
    result = round(delta.total_seconds() / 3600,2)
    if result<0:
        return 24+result
    else:
        return result

def billableHrsNew(timein, timeout, shift):
    is_flexi = False
    if isNaN(timein) or isNaN(timeout):
        return 0
    df_sched = db_sched[db_sched['Department'] == shift]
    df_sched = df_sched.reset_index()
    # print(df_sched)
    if df_sched.empty:
        # print("di ko kilala dept nya: {}".format(shift))
        df_sched = db_sched[db_sched['Department'] == "default"]
        df_sched = df_sched.reset_index()
    else:
        if(df_sched['Remarks'][0]=="flexi"):
            print("flexi boi")
            is_flexi = True

        
        # print("dept: {} || sched: {}-{} || timeinout: {}-{}".format(shift,df_sched['Day Shift In'][0],df_sched['Day Shift Out'][0],timein,timeout))
        # row['Gender']=df_new['Gender'][0]
        # row['Staff No'] = df_new['Staff No'][0]
    # return 0,0 # hrs, is_dayshift

    myin = timein.split(":")
    myout = timeout.split(":")
    inhr = int(myin[0])
    inmin = int(myin[1])
    outhr = int(myout[0])
    outmin = int(myout[1])

    hrs = outhr-inhr
    # # print("hrs:{}".format(hrs))
    if is_flexi:
        hrs = computeHrs(timein,timeout)
        if hrs>=11.75:
            return 12,1
        elif hrs<11.75 and hrs>=11:
            return 11,1
        if hrs>=6 and hrs<11:
            return 6,1
        else:
            return 0,1
    outputhrs = 0
    outputshift = 0
    late = 0
    if inhr>=(df_sched['Day Shift In'][0] - 2) and inhr<(df_sched['Day Shift In'][0] + 1):
        outputshift = 1 #am shift
        if inhr==(df_sched['Day Shift In'][0] + 1) and inmin>15 and inmin<=59:
            late = 1
        else:
            late = 0
    elif inhr>=(df_sched['Night Shift In'][0] - 2) and inhr<(df_sched['Night Shift In'][0] + 1):
        outputshift = 0 #pm shift
        if inhr==(df_sched['Night Shift In'][0] + 1) and inmin>15 and inmin<=59:
            late = 1
        else:
            late = 0
    else:
        return 0,0 #dont accept time ins that are outside the 5-8 and 17-20 window hrs
    # print("computing: inhr={} schedin={} outputshift={}".format(inhr,df_sched['Day Shift In'][0],outputshift))
    
    if outputshift==1:
        conv_shiftin = str(int(df_sched['Day Shift In'][0]))+":00:00"
    else:
        conv_shiftin = str(int(df_sched['Night Shift In'][0]))+":00:00"
    
    outdif = computeHrs(conv_shiftin,timeout)
    
    if outdif >=12:
        outputhrs = 12 - late
    elif outdif>=6 and outdif<12:
        outputhrs = 6 - late
    else:
        outputhrs = 0
    

    # if outputhrs != 12:#not (conv_shiftin == "7:00:00" or conv_shiftin == "19:00:00"):
    #     print("sched ({}): [{} {}] {} - {} = {}hrs = {}".format(shift,outputshift,timein,conv_shiftin,timeout,computeHrs(conv_shiftin,timeout),outputhrs))
    return outputhrs,outputshift

def billableHrs(timein, timeout, shift):
    if isNaN(timein) or isNaN(timeout):
        return 0
    if shift == "HOTEL SECURITY (DRIVER)":
        offset = -1
    elif shift == "HOTEL SECURITY (FOYER)": #3pm-3am
        offset = 8
    elif shift == "HOTEL SECURITY (TRAFFIC)" or shift =="ARMEDGUARD (TRAFFIC)": #12nn to 12 mn
        offset = 5
    elif shift == "ADMIN I.T":
        offset = 2
    else:
        offset = 0

    myin = timein.split(":")
    myout = timeout.split(":")
    inhr = int(myin[0])
    inmin = int(myin[1])
    outhr = int(myout[0])
    outmin = int(myout[1])
    hrs = outhr-inhr
    # print("hrs:{}".format(hrs))
    if offset==5:
        if outhr>=0 and outhr<=3:
            outhr = outhr+24
    if offset==8: 
        if outhr>=0 and outhr<=6:
            outhr = outhr+24
    if offset==2:
        hrs = computeHrs(timein,timeout)
        if hrs>=11.75:
            return 12
        elif hrs<11.75 and hrs>=11:
            return 11
        if hrs>=6 and hrs<11:
            return 6
        else:
            return 0
    else:
        if inhr>=5+offset and inhr<8+offset:
            shift = 1 #am shift
            if inhr==7+offset and inmin>15 and inmin<=59:
                late = 1
            else:
                late = 0
        elif inhr>=17+offset and inhr<20+offset:
            shift = 2 #pm shift
            if inhr==19+offset and inmin>15 and inmin<=59:
                late = 1
            else:
                late = 0
        else:
            return 0 #dont accept time ins that are outside the 5-8 and 17-20 window hrs
        
        if shift==1:
            if outhr>=19+offset and outhr<=23+offset:
                return 12-late
            elif outhr>=13+offset and outhr<19+offset:
                time=outhr-(7+offset)-late #halfday atleaest 6 hrs 
                if time >=6:
                    return 6
                else:
                    return 0
            else:
                return 0
        elif shift==2:
            if outhr >=7+offset and outhr<=11+offset:
                return 12-late
            elif outhr>=1+offset and outhr<7+offset:
                time=outhr-(19+offset)-late+24 #halfday atleaest 6 hrs 
                if time >=6:
                    return 6
                else:
                    return 0
            else:
                return 0



def openFile():
    if(tinputformat.get()==1):
        openTransactionalFile()
    else:
        openDailyComplete()

def openDailyComplete():
    global df
    global dfproc

    filename = fd.askopenfilename()
    
    hello = "file: " + filename
    myfilenamelabel.config(text=hello)
    statuslabel.config(text='status: Opening file...')
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
    # result = dfproc.head(10)
    # print("First 10 rows of the DataFrame:")
    print(dfproc.head(10))
    # listnames = (dfproc['Name'].unique())
    # for names in listnames:
    #     # print("ind:{} row:{}".format(index,cntr))
    #     # statustext = "status: Processing data...("+str(totcnt)+"/"+str(total)+")"
    #     # statuslabel.config(text=statustext)
    #     # root.update()
    #     # totcnt +=1
        
    #     df_byname = dfproc[dfproc['Name'] == names]
    #     df_byname = df_byname.reset_index()
    #     if df_byname

    statuslabel.config(text='status: File opened, ready to generate report')
    
    # print(dfproc)
    # print("list of guards")
    # testguy = 11
    # print(dfproc['Name'])
    # print("si {}:".format(testguy))
    # print(dfproc['Name'][testguy])
    # employeehrs = computeHrs(dfproc['Time In'][testguy],dfproc['Time Out'][testguy])
    # billablehrs = billableHrs(dfproc['Time In'][testguy],dfproc['Time Out'][testguy])
    # print("tothrs:{} billable:{}".format(employeehrs,billablehrs))

def openTransactionalFile():
    global df
    global dfproc

    filename = fd.askopenfilename()
    
    hello = "file: " + filename
    myfilenamelabel.config(text=hello)
    statuslabel.config(text='status: Opening file...')
    excellabel.config(text='')
    pdflabel.config(text='')
    # myLabel.pack(pady=1)
    # wb = openpyxl.load_workbook(filename)
    # ws = wb.active
    df = pd.read_excel(filename)

    # data = list(ws.iter_rows(values_only=True))
    print(df)
    print("one  that mattersz")
    # df.rename(columns=df.iloc[6])
    # df.columns = df.iloc[6]
    # df2 = df.drop(df.index[0:6])
    result = df.head(10)
    print("First 10 rows of the DataFrame:")
    print(result)
    
    # dfproc = df
    # dfproc.columns = df.iloc[0]
    print(df.columns)
    # dfproc = dfproc.loc[:,dfproc.columns.notna()]
    # dfproc = dfproc.dropna(axis=1)
    # dfproc = dfproc.reset_index()
    # dfproc = dfproc.drop('index', axis=1)
    df = df.sort_values(by=["Name","Date & Time"])
    df = df.reset_index()
    df = df.drop('index', axis=1) 
    df = df.drop('Company', axis=1) 
    df = df.drop('Vehicle No', axis=1) 
    # df['Date'] = str([d.date() for d in df['Date & Time']])
    # df['Time'] = str([d.time() for d in df['Date & Time']])
    df['Date'] = [d.strftime("%Y-%m-%d") for d in df['Date & Time']]
    df['Time'] = [d.strftime("%H:%M:%S") for d in df['Date & Time']]
    # df['Department'] = db[db['Name'] == row['Name']]
    df['Department'] = " "
    prevname = ""
    prevtrans = ""
    prevdate = ""
    total = len(df.index)
    for index,row in df.iterrows():
        # if index >=200:
        #     break
        # print(df.loc[[index]])
        statustext = "status: Processing transactional report...("+str(index)+"/"+str(total)+")"
        statuslabel.config(text=statustext)
        root.update()
        if prevname != row['Name']:
            # new data row
            temp =db.loc[db['Name'] == row['Name'], 'Department']
            if(len(temp)>=1):
                # print("temp: {}".format(temp.item())).
                mydept = str(temp.item())
                df.loc[index,'Department']= mydept
            else:
                mydept = ""

            apnd_newrow = pd.DataFrame({'Date':row['Date'],'Department':mydept,'Card No':row['Card No'],'Name':row['Name'],'Staff No':row['Staff No']},index=[0])
            if row['Transaction'] == "Valid Entry Access":
                apnd_newrow['Time In'] = row['Time']
            if row['Transaction'] == "Valid Exit Access": #if exit
                apnd_newrow['Time Out'] = row['Time']
            dfproc = dfproc.append(apnd_newrow,ignore_index=True)
        else: 
            if prevtrans != row['Transaction']:
                if row['Transaction'] == "Valid Entry Access":
                    apnd_newrow = pd.DataFrame({'Date':row['Date'],'Name':row['Name'],'Time In':row['Time'],'Department':mydept,'Card No':row['Card No'],'Staff No':row['Staff No']},index=[0])
                    dfproc = dfproc.append(apnd_newrow,ignore_index=True)
                if row['Transaction'] == "Valid Exit Access": #if exit
                    dfproc.loc[dfproc.index[-1], 'Time Out'] = row['Time']
            else:
                if row['Transaction'] == "Valid Entry Access":
                    if row['Date'] != prevdate:
                        # new data row
                        pass
                    else:
                        pass
                if row['Transaction'] == "Valid Exit Access": #if exit
                    # 
                    if row['Date'] != prevdate:
                        # new data row
                        pass
                    else:
                        pass

        # aggr_newrow = pd.DataFrame({'Name':index},index=[0])
        # aggr_newrow['Total'] = int(aggr_newrow['Total Hrs'])
        # dfproc = dfproc.append(aggr_newrow,ignore_index=True)
        # dfproca.loc[dfproc.index[-1], 'a'] = 4.0

        #if new name                            : new row matic
        #if old name, trans  is exit to entry   : new row
        #if old name, old trans, 


        prevname = row['Name']
        prevtrans = row['Transaction']
        prevdate = row['Date']
    # df['Date'] = " "
    # df['Time'] = " "
    # for index,row in df.iterrows():
    #     mydate  = datetime.strptime(str(row['Date & Time']),'%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
    #     df['Date'][index]  = mydate
    #     mytime = datetime.strptime(str(row['Date & Time']),'%Y-%m-%d %H:%M:%S').strftime('%H:%M:%S')
    #     df['Time'][index] = mytime
    # result = dfproc.head(10)
    # print("First 10 rows of the DataFrame:")
    print(dfproc.head(60))
    # listnames = (dfproc['Name'].unique())
    # for names in listnames:
    #     # print("ind:{} row:{}".format(index,cntr))
    #     # statustext = "status: Processing data...("+str(totcnt)+"/"+str(total)+")"
    #     # statuslabel.config(text=statustext)
    #     # root.update()
    #     # totcnt +=1
        
    #     df_byname = dfproc[dfproc['Name'] == names]
    #     df_byname = df_byname.reset_index()
    #     if df_byname

    statuslabel.config(text='status: File opened, ready to generate report')
    
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
        if (bimoformat.get()):
            generateNewBiMonthlyReport()
        else:
            generateBiMonthlyReport()
    if selection ==2:
        generatePEZAReport()

def generateNewBiMonthlyReport():
    global dfproc,savexlsx,savepdf
    
    
    print(dfproc)
    print("generating...")
    statuslabel.config(text='status: generating...')
    dfproc['Total Hours'] = " "
    dfproc['Billable Hours'] = " "
    dfproc['is_DS'] = " "
    # for index, row in dfproc.iterrows():
    #     myhrs = computeHrs(row['Time In'],row['Time Out'])
    #     billablehrsnew = billableHrsNew(row['Time In'],row['Time Out'],row['Department'])
    #     row['Total Hours']=myhrs
    #     row['Billable Hours']=billablehrsnew[0]
    #     row['is_DS'] = billablehrsnew[1]

    for index, row in dfproc.iterrows():
        
        # print("index:{} row:{}".format(index,row))
        if isNaN(row['Time In']) or isNaN(row['Time Out']):
            # listtodrop.append(index)
            pass
        else:
            myhrs = computeHrs(row['Time In'],row['Time Out'])
            billablehrsnew = billableHrsNew(row['Time In'],row['Time Out'],row['Department'])
            row['Total Hours']=myhrs
            # try:
            row['Billable Hours']=billablehrsnew[0]
            row['is_DS'] = billablehrsnew[1]
            # except TypeError:
            #     print(row)
        
            # continue
    
    dfproc = dfproc.reset_index()
    dfproc = dfproc.drop('index', axis=1) 
    print(dfproc)
    col = dfproc.pop("Name")
    dfproc.insert(0,col.name,col)
    print("moving aroundd")

    dfproc['Day'] = " "
    listtodrop = []
    for index, row in dfproc.iterrows():
        # df_new = db[db['Card N°.'] == row['Card No']]
        df_new = db[db['Name'] == row['Name']]
        df_new = df_new.reset_index()

        myday = datetime.strptime(row['Date'],'%Y-%m-%d').strftime('%a')
        row['Day']=myday
        # print("index:{} row:{}".format(index,row))
        # myhrs = computeHrs(row['Time In'],row['Time Out'])
        # billablehrs = billableHrs(row['Time In'],row['Time Out'])
        # row['Total Hours']=myhrs
        # row['Billable Hours']=billablehrs
        if isNaN(row['Time In']) or isNaN(row['Time Out']) or row['Time In']==' 'or row['Time Out']==' ':
            listtodrop.append(index)
            # continue

    col = dfproc.pop("Department")
    dfproc.insert(3,"Category",col)
    col = dfproc.pop("Day")
    dfproc.insert(5,col.name,col)
    dfproc = dfproc.fillna('')
    print(dfproc)

    listnames = (dfproc['Name'].unique())
    # print(listnames)
    # dfmonth = pd.DataFrame(data='',columns=dfproc.columns)
    dfmonth = dfproc.iloc[:0]
    print(dfmonth)
    # dfmonth = pd.DataFrame
    month_sel = monthvariable.get()
    cutoff_sel = cutoffvariable.get()
    year_sel = yearvariable.get()
    print("mo:{} cutoff:{} yr:{}".format(month_sel,cutoff_sel,year_sel))
    if cutoff_sel =="1-15":
        days = ["01","02","03","04","05","06","07","08","09","10","11","12","13","14","15"]
    else:
        days = ["16","17","18","19","20","21","22","23","24","25","26","27","28"]#,"29","30","31"]
        if month_sel in month_30:
            days.extend(["29","30"])
        elif month_sel in month_31:
            days.extend(["29","30","31"])
        elif month_sel=="Feb":
            if ((int(year_sel) % 400 == 0) or (int(year_sel) % 100 != 0) and (int(year_sel) % 4 == 0)):
                days.extend(["29"])
    print(days)
    # df_aggregated = pd.DataFrame
    # df_aggregated["Name"]=""
    aggr_columns = ["Name"]
    for items in days:
        title = str(items) + " ds"
        aggr_columns.append(title)
        title = str(items) + " ns"
        aggr_columns.append(title)
    aggr_columns.append("Total Hrs")
    aggr_columns.append("Total Days")
    aggr_columns.append("Total DS Hrs")
    aggr_columns.append("Total NS Hrs")
    aggr_columns.append("Total")
    df_aggregated = pd.DataFrame(columns=aggr_columns)
    print(df_aggregated)
    total = len(listnames)
    totcnt = 0
    for index in listnames:
        # print("ind:{} row:{}".format(index,cntr))
        statustext = "status: Processing data...("+str(totcnt)+"/"+str(total)+")"
        statuslabel.config(text=statustext)
        root.update()
        totcnt +=1
        
        df_byname = dfproc[dfproc['Name'] == index]
        df_byname = df_byname.reset_index()
        aggr_newrow = pd.DataFrame({'Name':index},index=[0])
        # df_byname = df[df['EPS'].notna()]
        # print("df_byname below")
        # print(df_byname)
        runningtotal=0
        totalhrs = 0
        totaldays = 0
        totalhrs_ds = 0
        totalhrs_ns = 0
        for x in days :
            mymonth = datetime.strptime(month_sel,"%b").strftime('%m')
            date = year_sel+"-"+mymonth + "-" +x
            # print(date)
            new_row = pd.DataFrame({'Name':index,'Date':date},index=[0])
            for cntr, row in df_byname.iterrows():
                df_date = df_byname[df_byname['Date'] == date]
                df_date = df_date.reset_index()
                if not df_date.empty:
                    if len(df_date.index)>=1:
                        new_row['Time In']=df_date['Time In'][0]
                        new_row['Time Out']=df_date['Time Out'][0]
                        new_row['Total Hours']=df_date['Total Hours'][0]
                        new_row['Billable Hours']=df_date['Billable Hours'][0]
                        new_row['Category']=df_date['Category'][0]
                        runningtotal = runningtotal + int(new_row['Billable Hours'])
                        
                        if int(df_date['is_DS'])==1:
                            aggr_newrow[str(x)+" ds"] = df_date['Billable Hours'][0]
                            totalhrs_ds = totalhrs_ds + int(df_date['Billable Hours'][0])
                        elif int(df_date['is_DS'])==0:
                            aggr_newrow[str(x)+" ns"] = df_date['Billable Hours'][0]
                            totalhrs_ns = totalhrs_ns + int(df_date['Billable Hours'][0])
                        totaldays = totaldays + 1

                        break
                    # if len(df_date.index)>1:
                    #     if (not isNan(row['Time In'])) and (not isNan(row['Time Out'])):
                    #         new_row['Time In']=df_date['Time In'][0]
                    #         new_row['Time Out']=df_date['Time Out'][0]
                    #         new_row['Total Hours']=df_date['Total Hours'][0]
                    #         new_row['Billable Hours']=df_date['Billable Hours'][0]
                    #         new_row['Employee Number'] = df_date['Employee Number'][0]
                    #         runningtotal = runningtotal + int(new_row['Billable Hours'])

            myday = datetime.strptime(date,'%Y-%m-%d').strftime('%a')
            new_row['Date'] = datetime.strptime(date,'%Y-%m-%d').strftime('%b %d, %Y')
            new_row['Day']=myday
            # runningtotal = runningtotal + new_row['Billable Hours']
            
            # dfmonth = pd.concat([new_row,dfproc.loc[:]]).reset_index(drop=True)
            dfmonth = dfmonth.append(new_row,ignore_index=True)
        aggr_newrow['Total Hrs'] = totalhrs_ds + totalhrs_ns
        aggr_newrow['Total Days'] = totaldays
        aggr_newrow['Total DS Hrs'] = totalhrs_ds
        aggr_newrow['Total NS Hrs'] = totalhrs_ns
        aggr_newrow['Total'] = int(aggr_newrow['Total Hrs'])
        df_aggregated = df_aggregated.append(aggr_newrow,ignore_index=True)
        # new_row = pd.DataFrame({'Name':' ','Total Hours':"Total:", 'Billable Hours':runningtotal},index=[0])
        # dfmonth = dfmonth.append(new_row,ignore_index=True)
    # print("change date")
    # for index, row in dfmonth.iterrows():
    #     row['Date'] = datetime.strptime(str(row['Date']),'%Y-%m-%d').strftime('%b %d, %Y')
    
    dfmonth = dfmonth.fillna('')
    df_aggregated = df_aggregated.fillna('')

    # print(dfmonth.head(60))
    print(df_aggregated)
    statuslabel.config(text='status: Generating reports...')
    myfilenamelabel.config(text='file: ')
    if savepdf.get()==1:
        pdflabel.config(text='pdf: generating...')
        savePdflandscape(df_aggregated,3)
        # savePdf(df_aggregated,3)
    if savexlsx.get()==1:
        excellabel.config(text='spreadsheet: generating...')
        saveXlsx(df_aggregated,3)
    
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

def generateBiMonthlyReport():
    global dfproc,savexlsx,savepdf
    
    
    print(dfproc)
    print("generating...")
    statuslabel.config(text='status: generating...')
    dfproc['Total Hours'] = " "
    dfproc['Billable Hours'] = " "
    # for index, row in dfproc.iterrows():
    #     myhrs = computeHrs(row['Time In'],row['Time Out'])
    #     billablehrs = billableHrsNew(row['Time In'],row['Time Out'],row['Department'])
    #     row['Total Hours']=myhrs
    #     row['Billable Hours']=billablehrs[0]
    #     # row['is_DS'] = billablehrs[1]
    
    for index, row in dfproc.iterrows():
        
        # print("index:{} row:{}".format(index,row))
        if isNaN(row['Time In']) or isNaN(row['Time Out']):
            # listtodrop.append(index)
            pass
        else:
            myhrs = computeHrs(row['Time In'],row['Time Out'])
            billablehrsnew = billableHrsNew(row['Time In'],row['Time Out'],row['Department'])
            row['Total Hours']=myhrs
            # try:
            row['Billable Hours']=billablehrsnew[0]
            # except TypeError:
            #     print(row)
        
            # continue

    
    dfproc = dfproc.reset_index()
    dfproc = dfproc.drop('index', axis=1) 
    print(dfproc)
    col = dfproc.pop("Name")
    dfproc.insert(0,col.name,col)
    print("moving aroundd")
    dfproc['Gender'] = " "
    dfproc['Remarks'] = " "
    dfproc['Day'] = " "
    listtodrop = []
    for index, row in dfproc.iterrows():
        # df_new = db[db['Card N°.'] == row['Card No']]
        df_new = db[db['Name'] == row['Name']]
        df_new = df_new.reset_index()
        if df_new.empty:
            row['Gender'] = " "
            row['Staff No'] = ""
        else:
            row['Gender']=df_new['Gender'][0]
            row['Staff No'] = df_new['Staff No'][0]
        myday = datetime.strptime(row['Date'],'%Y-%m-%d').strftime('%a')
        row['Day']=myday
        # print("index:{} row:{}".format(index,row))
        # myhrs = computeHrs(row['Time In'],row['Time Out'])
        # billablehrs = billableHrs(row['Time In'],row['Time Out'])
        # row['Total Hours']=myhrs
        # row['Billable Hours']=billablehrs
        if isNaN(row['Time In']) or isNaN(row['Time Out']) or row['Time In']==' ' or row['Time Out']==' ':
            listtodrop.append(index)
            # continue
    # print(listtodrop)
    # dfproc = dfproc.drop(listtodrop)

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

    listnames = (dfproc['Name'].unique())
    # print(listnames)
    # dfmonth = pd.DataFrame(data='',columns=dfproc.columns)
    dfmonth = dfproc.iloc[:0]
    print(dfmonth)
    # dfmonth = pd.DataFrame
    month_sel = monthvariable.get()
    cutoff_sel = cutoffvariable.get()
    year_sel = yearvariable.get()
    print("mo:{} cutoff:{} yr:{}".format(month_sel,cutoff_sel,year_sel))
    if cutoff_sel =="1-15":
        days = ["01","02","03","04","05","06","07","08","09","10","11","12","13","14","15"]
    else:
        days = ["16","17","18","19","20","21","22","23","24","25","26","27","28"]#,"29","30","31"]
        if month_sel in month_30:
            days.extend(["29","30"])
        elif month_sel in month_31:
            days.extend(["29","30","31"])
        elif month_sel=="Feb":
            if ((int(year_sel) % 400 == 0) or (int(year_sel) % 100 != 0) and (int(year_sel) % 4 == 0)):
                days.extend(["29"])
    print(days)
    total = len(listnames)
    totcnt = 0
    for index in listnames:
        # print("ind:{} row:{}".format(index,cntr))
        statustext = "status: Processing data...("+str(totcnt)+"/"+str(total)+")"
        statuslabel.config(text=statustext)
        root.update()
        totcnt +=1
        
        df_byname = dfproc[dfproc['Name'] == index]
        df_byname = df_byname.reset_index()
        # df_byname = df[df['EPS'].notna()]
        # print(df_byname)
        runningtotal=0
        for x in days :
            mymonth = datetime.strptime(month_sel,"%b").strftime('%m')
            date = year_sel+"-"+mymonth + "-" +x
            # print(date)
            new_row = pd.DataFrame({'Name':index,'Date':date},index=[0])
            for cntr, row in df_byname.iterrows():
                df_date = df_byname[df_byname['Date'] == date]
                df_date = df_date.reset_index()
                if not df_date.empty:
                    if len(df_date.index)>=1:
                        new_row['Time In']=df_date['Time In'][0]
                        new_row['Time Out']=df_date['Time Out'][0]
                        new_row['Total Hours']=df_date['Total Hours'][0]
                        new_row['Billable Hours']=df_date['Billable Hours'][0]
                        new_row['Employee Number'] = df_date['Employee Number'][0]
                        new_row['Gender']=df_date['Gender'][0]
                        new_row['Category']=df_date['Category'][0]
                        new_row['Staff No']=df_date['Staff No'][0]
                        runningtotal = runningtotal + int(float(new_row['Billable Hours']))
                        break
                    # if len(df_date.index)>1:
                    #     if (not isNan(row['Time In'])) and (not isNan(row['Time Out'])):
                    #         new_row['Time In']=df_date['Time In'][0]
                    #         new_row['Time Out']=df_date['Time Out'][0]
                    #         new_row['Total Hours']=df_date['Total Hours'][0]
                    #         new_row['Billable Hours']=df_date['Billable Hours'][0]
                    #         new_row['Employee Number'] = df_date['Employee Number'][0]
                    #         runningtotal = runningtotal + int(new_row['Billable Hours'])

            myday = datetime.strptime(date,'%Y-%m-%d').strftime('%a')
            new_row['Date'] = datetime.strptime(date,'%Y-%m-%d').strftime('%b %d, %Y')
            new_row['Day']=myday
            # runningtotal = runningtotal + new_row['Billable Hours']
            
            # dfmonth = pd.concat([new_row,dfproc.loc[:]]).reset_index(drop=True)
            dfmonth = dfmonth.append(new_row,ignore_index=True)
        new_row = pd.DataFrame({'Name':' ','Total Hours':"Total:", 'Billable Hours':runningtotal},index=[0])
        dfmonth = dfmonth.append(new_row,ignore_index=True)
        new_row = pd.DataFrame({'Name':' '},index=[0])
        # new_row = pd.DataFrame(
        #         {'Name':' ',
        #         'Employee Number':"-----", 
        #         'Gender':"-----", 
        #         'Category':"-----", 
        #         'Date':"-----", 
        #         'Day':"-----", 
        #         'Time In':"-----", 
        #         'Time Out':"-----", 
        #         'Total Hours':"-----", 
        #         'Billable Hours':"-----", 
        #         'Remarks':"--------"},index=[0])
        dfmonth = dfmonth.append(new_row,ignore_index=True)
    # print("change date")
    # for index, row in dfmonth.iterrows():
    #     row['Date'] = datetime.strptime(str(row['Date']),'%Y-%m-%d').strftime('%b %d, %Y')
    
    dfmonth = dfmonth.fillna('')
    col = dfmonth.pop("Employee Number")
    col = dfmonth.pop("Staff No")
    dfmonth.insert(1,"Employee Number",col)
    print(dfmonth)
    statuslabel.config(text='status: Generating reports...')
    myfilenamelabel.config(text='file: ')
    if savepdf.get()==1:
        pdflabel.config(text='pdf: generating...')
        savePdf(dfmonth,1)
    if savexlsx.get()==1:
        excellabel.config(text='spreadsheet: generating...')
        saveXlsx(dfmonth,1)
    
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

def generatePEZAReport():
    global dfproc,savexlsx,savepdf
    # dfproc=db
    dfproc.iloc[0:0]
    dfproc = db[['Name', 'Gender', 'Staff No', 'Department']].copy()
    dfproc = dfproc.drop([0])
    dfproc = dfproc.reset_index()
    dfproc = dfproc.drop('index', axis=1) 
    dfproc = dfproc.sort_values(by=["Name"])
    # print(dfproc)
    # dfproc['USER ID'] = dfproc['USER ID'].fillna(0).astype('int')
    print(dfproc)
    statuslabel.config(text='status: PEZA Report generated')
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
    listtodrop = []
    for index, row in dfproc.iterrows():
        
        # print("index:{} row:{}".format(index,row))
        if isNaN(row['Time In']) or isNaN(row['Time Out']):
            listtodrop.append(index)
        else:
            myhrs = computeHrs(row['Time In'],row['Time Out'])
            billablehrsnew = billableHrsNew(row['Time In'],row['Time Out'],row['Department'])
            row['Total Hours']=myhrs
            try:
                row['Billable Hours']=billablehrsnew[0]
            except TypeError:
                print(row)
        
            # continue
    # print(listtodrop)
    # dfproc = dfproc.drop(listtodrop)

    dfproc = dfproc.reset_index()
    dfproc = dfproc.drop('index', axis=1) 
    print(dfproc)
    col = dfproc.pop("Name")
    dfproc.insert(0,col.name,col)
    print("moving aroundd")
    dfproc['Gender'] = " "
    dfproc['Remarks'] = " "
    dfproc['Day'] = " "
    total = len(dfproc.index)
    totcnt = 0
    for index, row in dfproc.iterrows():
        statustext = "status: Processing data...("+str(totcnt)+"/"+str(total)+")"
        statuslabel.config(text=statustext)
        root.update()
        totcnt+=1
        # df_new = db[db['Card N°.'] == row['Card No']]
        df_new = db[db['Name'] == row['Name']]
        df_new = df_new.reset_index()
        if df_new.empty:
            row['Gender'] = " "
            row['Staff No'] = ""
        else:
            row['Gender']=df_new['Gender'][0]
            row['Staff No'] = df_new['Staff No'][0]
            # print("meron naman")
        myday = datetime.strptime(row['Date'],'%Y-%m-%d').strftime('%a')
        row['Date'] = datetime.strptime(row['Date'],'%Y-%m-%d').strftime('%b %d, %Y')
        row['Day']=myday
    col = dfproc.pop("Card No")
    #dfproc.insert(1,"Employee Number",col)
    col = dfproc.pop("Gender")
    dfproc.insert(1,col.name,col)
    col = dfproc.pop("Department")
    dfproc.insert(2,"Category",col)
    col = dfproc.pop("Day")
    dfproc.insert(4,col.name,col)
    col = dfproc.pop("Staff No")
    dfproc.insert(1,"Employee Number",col)
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
    month_sel = monthvariable.get()
    cutoff_sel = cutoffvariable.get()
    year_sel = yearvariable.get()
    if reporttype==0:
        initfile = "SummaryDailyReport_"+now.strftime("%d-%m-%Y")
    elif reporttype==1:
        initfile = "SummaryBiMonthlyReport_"+month_sel+cutoff_sel+","+year_sel
    elif reporttype==2:
        initfile = "SummaryPEZAReport_"+now.strftime("%d-%m-%Y")
    elif reporttype==3:
        initfile = "SummaryAggregatedBiMonthlyReport_"+month_sel+cutoff_sel+","+year_sel
    filename_out=fd.asksaveasfile(mode='w',defaultextension=".xlsx",initialfile=initfile,filetypes=(("Excel Workbook", "*.xlsx"),("All Files", "*.*")))
    print(filename_out.name)
    statustext = "status: Creating spreadsheet..."
    statuslabel.config(text=statustext)
    writer = pd.ExcelWriter(filename_out.name, engine='xlsxwriter')
    # dfproc = dfproc.style.applymap(lambda x: "border-style: hair; border-color: black")
    if reporttype==3:
        dfproc.to_excel(writer,startrow=1,index=False, sheet_name='Sheet1')
        # worksheet = writer.sheets['Sheet1']
        title = "PERIOD: "+month_sel+" "+cutoff_sel+", "+year_sel
        writer.sheets['Sheet1'].write(0,0,title)
        # writer.save()
    else:
        dfproc.to_excel(writer,index=False, sheet_name='Sheet1')
    for column in dfproc:
        column_width = max(dfproc[column].astype(str).map(len).max(), len(column))
        col_idx = dfproc.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

    writer.save()
    statustext = "status: Created spreadsheet..."
    statuslabel.config(text=statustext)
    excellabel.config(text='spreadsheet: '+filename_out.name)

def savePdf(dfproc,reporttype):

    now = datetime.now()
    if reporttype==0:
        initfile = "SummaryDailyReport_"+now.strftime("%d-%m-%Y")
        figtitle = "Kenichi Security (Okada)\nDaily Time & Attendance "+now.strftime("%d-%m-%Y")+"\n"
    elif reporttype==1:
        initfile = "SummaryBiMonthlyReport_"+now.strftime("%d-%m-%Y")
        month_sel = monthvariable.get()
        cutoff_sel = cutoffvariable.get()
        year_sel = yearvariable.get()
        figtitle = "Kenichi Security (Okada)\nBimonthly Billing Covering "+month_sel
        # print("mo:{} cutoff:{} yr:{}".format(month_sel,cutoff_sel,year_sel))
        if cutoff_sel =="1-15":
            figtitle += " 1-15, "+year_sel+"\n"
        else:
            if month_sel in month_30:
                figtitle += " 16-30, "+year_sel+"\n"
            elif month_sel in month_31:
                figtitle += " 16-31, "+year_sel+"\n"
            elif month_sel=="Feb":
                if ((int(year_sel) % 400 == 0) or (int(year_sel) % 100 != 0) and (int(year_sel) % 4 == 0)):
                    figtitle += " 16-29, "+year_sel+"\n"
                else:
                    figtitle += " 16-28, "+year_sel+"\n"
    elif reporttype==2:
        initfile = "SummaryPEZAReport_"+now.strftime("%d-%m-%Y")
        figtitle = "Kenichi Security (Okada)\nPEZA Report as of "+now.strftime("%d-%m-%Y")+"\n"
    filename_out=fd.asksaveasfile(mode='w',defaultextension=".pdf",initialfile=initfile,filetypes=(("PDF file", "*.pdf"),("All Files", "*.*")))
    print(filename_out.name)
    statustext = "status: Creating PDF..."
    statuslabel.config(text=statustext)
    dfproc = dfproc.rename({'USER ID':'User ID','Staff No':'Employee\nNumber','Employee Number':'Employee\nNumber', 'Total Hours':'Total\nHours', 'Billable Hours':'Billable\nHours'},axis=1)
    groups = dfproc.groupby(np.arange(len(dfproc.index))//50)
    pp = PdfPages(filename_out.name)
    
    for (frameno, frame) in groups:    
        statustext = "status: Creating PDF...("+str(frameno)+"/"+str(groups.ngroups)+")"
        statuslabel.config(text=statustext)
        root.update()
        fig, ax =plt.subplots(figsize=(8.5,11))
        ax.set_title(figtitle,fontsize=10)
        ax.axis('tight')
        ax.axis('off')
        if reporttype==2:
            the_table = ax.table(cellText=frame.values,colLabels=frame.columns,cellLoc='left',loc='center',colWidths=[.25,.076,.12,.22])
        else:
            the_table = ax.table(cellText=frame.values,colLabels=frame.columns,cellLoc='left',loc='center',colWidths=[.25,.12,.076,.22,.11,.046,.09,.09,.052,.052,.072])
        cellDict = the_table.get_celld()
        for i in range(0,len(frame.columns)):
            cellDict[(0,i)].set_height(.03)
            for j in range(1,len(frame.index)+1):
                cellDict[(j,i)].set_height(.02)

        the_table.auto_set_font_size(False)
        the_table.set_fontsize(6)
        # the_table.auto_set_column_width(col=[0,1,2,3,4])
        plt.close()
        # the_table[0, col]._text.set_horizontalalignment('left') 
        
        pp.savefig(fig, bbox_inches='tight')
    pp.close()
    pdflabel.config(text='pdf: '+filename_out.name)
    statustext = "status: Created PDF..."
    statuslabel.config(text=statustext)
    # os.startfile(filename_out.name)
# def saveFile():
    # print("saving...")
    # html_string = df.to_html()
    # pdfkit.from_string(html_string,"output_file.pdf")
    # print("file saved")

def savePdflandscape(dfproc,reporttype):
    now = datetime.now()
    month_sel = monthvariable.get()
    cutoff_sel = cutoffvariable.get()
    year_sel = yearvariable.get()

    initfile = "SummaryAggregatedBiMonthlyReport_"+month_sel
    figtitle = "Kenichi Security (Okada)\nAggregated Bimonthly Billing Covering "+month_sel
    # print("mo:{} cutoff:{} yr:{}".format(month_sel,cutoff_sel,year_sel))
    if cutoff_sel =="1-15":
        figtitle += " 1-15, "+year_sel+"\n"
        initfile += "1-15,"+year_sel
        numofdays = 15
    else:
        if month_sel in month_30:
            figtitle += " 16-30, "+year_sel+"\n"
            initfile += "16-30,"+year_sel
            numofdays = 15
        elif month_sel in month_31:
            figtitle += " 16-31, "+year_sel+"\n"
            initfile += "16-31,"+year_sel
            numofdays = 16
        elif month_sel=="Feb":
            if ((int(year_sel) % 400 == 0) or (int(year_sel) % 100 != 0) and (int(year_sel) % 4 == 0)):
                figtitle += " 16-29, "+year_sel+"\n"
                initfile += "16-29,"+year_sel
                numofdays = 14
            else:
                figtitle += " 16-28, "+year_sel+"\n"
                initfile += "16-28,"+year_sel
                numofdays = 13
    filename_out=fd.asksaveasfile(mode='w',defaultextension=".pdf",initialfile=initfile,filetypes=(("PDF file", "*.pdf"),("All Files", "*.*")))
    print(filename_out.name)
    statustext = "status: Creating PDF..."
    statuslabel.config(text=statustext)
    dfproc = dfproc.rename({'USER ID':'User ID','Staff No':'Employee\nNumber','Employee Number':'Employee\nNumber', 'Total Hours':'Total\nHours', 'Billable Hours':'Billable\nHours'},axis=1)
    groups = dfproc.groupby(np.arange(len(dfproc.index))//50)
    pp = PdfPages(filename_out.name)
    
    for (frameno, frame) in groups:    
        statustext = "status: Creating PDF...("+str(frameno)+"/"+str(groups.ngroups)+")"
        statuslabel.config(text=statustext)
        root.update()
        fig, ax =plt.subplots(figsize=(11,8.5))
        ax.set_title(figtitle,fontsize=10)
        ax.axis('tight')
        ax.axis('off')
        # if numofdays==13:
        #     mycolwidth=[]
        # elif numofdays==14:
        #     mycolwidth=[]
        # elif numofdays==15:
        #     mycolwidth=[.18,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.025,.045,.045,.07,.07,.04]
        #     data=0
        #     for items in mycolwidth:
        #         data = data + items
        #     print(data)
        # else:
        #     mycolwidth=[]
        mycolwidth=[.18]
        cntr = 0
        for x in range(0,numofdays*2):
            mycolwidth.append(.025)
            cntr = cntr+1
        # print(/cntr)
        mycolwidth.append(.045)
        mycolwidth.append(.045)
        mycolwidth.append(.06)
        mycolwidth.append(.06)
        mycolwidth.append(.04)
        the_table = ax.table(cellText=frame.values,colLabels=frame.columns,cellLoc='left',loc='center',colWidths=mycolwidth)

        cellDict = the_table.get_celld()
        for i in range(0,len(frame.columns)):
            cellDict[(0,i)].set_height(.03)
            for j in range(1,len(frame.index)+1):
                cellDict[(j,i)].set_height(.02)

        the_table.auto_set_font_size(False)
        the_table.set_fontsize(5)
        # the_table.auto_set_column_width(col=[0,1,2,3,4])
        plt.close()
        # the_table[0, col]._text.set_horizontalalignment('left') 
        
        pp.savefig(fig, bbox_inches='tight')
    pp.close()
    pdflabel.config(text='pdf: '+filename_out.name)
    statustext = "status: Created PDF..."
    statuslabel.config(text=statustext)
    # os.startfile(filename_out.name)
# def saveFile():
    # print("saving...")
    # html_string = df.to_html()
    # pdfkit.from_string(html_string,"output_file.pdf")
    # print("file saved")

def print_selection():
    print("as pdf:{} as xlsx:{}".format(savepdf.get(),savexlsx.get()))
def print_selection1():
    print("bimonthly format:{}".format(bimoformat.get()))

# e = tk.Entry(root, width=50, font=('Helvetica',20)).pack(padx=10, pady=10)

# myButton = tk.Button(root, text="Enter naym",command=myClick)
# myButton.pack(padx=20)
headerlabel = tk.Label(root, text="Kenichi Security (Okada)",wraplength=400, justify="left",font=('Helvetica',15)).pack(side=tk.TOP,anchor=tk.NW,padx=20,pady=20)
tk.ttk.Separator(root, orient='horizontal').pack(fill='x',pady=5)
myButton2 = tk.Button(root, text="Open File",command=openFile,font=('Helvetica',13))
myButton2.pack(side=tk.TOP, anchor=tk.NW,padx=20,pady=10)
input_format = tk.Checkbutton(root, text='Transactional input file',variable=tinputformat, onvalue=1, offvalue=0, command=print_selection1,font=('Helvetica',9))
input_format.pack(anchor=tk.W,padx=20)
myfilenamelabel = tk.Label(root, text="file:",wraplength=400)
myfilenamelabel.pack(side=tk.TOP,anchor=tk.NW,padx=20)
tk.ttk.Separator(root, orient='horizontal').pack(fill='x',pady=5)





def ShowChoice():
    print(radbut.get())

radbut = tk.IntVar()
radbut.set(0)
radbut_values = [("Daily T&A Report", 0),
   	            ("Bimonthly Billing Report", 1),
                ("PEZA Report", 2)]
for text, value in radbut_values:
    tk.Radiobutton(root, 
                   text=text,
                   padx = 20, 
                   variable=radbut, 
                   command=ShowChoice,
                   value=value,font=('Helvetica',13)).pack(anchor=tk.W)
    if value == 1:
        month_drop = tk.OptionMenu(root, monthvariable, *dropdown_month).pack()#side=tk.LEFT,padx=20)
        cutoff_drop = tk.OptionMenu(root, cutoffvariable, *dropdown_cutoff).pack()#side=tk.LEFT,padx=20)

        year_entry.pack()
        bimo_format = tk.Checkbutton(root, text='New Format',variable=bimoformat, onvalue=1, offvalue=0, command=print_selection1,font=('Helvetica',10))
        bimo_format.pack()#side=tk.TOP,anchor=tk.W,padx=20)



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
root.wm_protocol("WM_DELETE_WINDOW", destroyer)
