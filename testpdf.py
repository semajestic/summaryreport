import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os,sys



def savePdf(dfproc,reporttype):

    # now = datetime.now()
    # if reporttype==0:
    #     initfile = "SummaryDailyReport_"+now.strftime("%d-%m-%Y")
    # elif reporttype==1:
    #     initfile = "SummaryBiMonthlyReport_"+now.strftime("%d-%m-%Y")
    # elif reporttype==2:
    #     initfile = "SummaryCEZAReport_"+now.strftime("%d-%m-%Y")
    filename_out="testpdf.pdf"
    print(filename_out)
    # statustext = "status: Creating PDF..."
    # statuslabel.config(text=statustext)
    
    print(np.arange(len(dfproc.index))//50)
    groups = dfproc.groupby(np.arange(len(dfproc.index))//50)
    print(groups.ngroups)
    pp = PdfPages(filename_out)
    for (frameno, frame) in groups:    
        # frame.to_csv("%s.csv" % frameno)
        fig, ax =plt.subplots(figsize=(8.5,11))
        # ax.suptitle('Kenichi Security (Okada)')
        ax.set_title("Kenichi Security (Okada)\nBi-monthly BillingReport (date1) to (date2)\n",fontsize=10)
        # plt.title('Mean WRFv3.5 LHF\n', fontsize=20)
        
        ax.axis('tight')
        ax.axis('off')
        the_table = ax.table(cellText=frame.values,colLabels=frame.columns,cellLoc='left',loc='center')
        the_table.auto_set_font_size(False)
        the_table.set_fontsize(6)
        the_table.auto_set_column_width(col=[0,1,2,3,4])

        cellDict = the_table.get_celld()
        for i in range(0,len(frame.columns)):
            cellDict[(0,i)].set_height(.03)
            for j in range(1,len(frame.index)+1):
                cellDict[(j,i)].set_height(.02)
        # the_table[0, col]._text.set_horizontalalignment('left') 
        
        pp.savefig(fig, bbox_inches='tight')
    pp.close()

    # pdflabel.config(text='pdf: '+filename_out)
    print("created")
    os.startfile(filename_out)
    # os.startfile(filename_out.name)
# def saveFile():
    # print("saving...")
    # html_string = df.to_html()
    # pdfkit.from_string(html_string,"output_file.pdf")
    # print("file saved")

df = pd.DataFrame(data=np.random.rand(280, 3), columns=list('ABC'))

savePdf(df,0)
# df = pd.DataFrame(data=np.random.rand(100, 3), columns=list('ABC'))

# savePdf(df,0)