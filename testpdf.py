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
    print(len(dfproc.index)%50)
    print(groups.ngroups)
    pp = PdfPages(filename_out)
    printsigs = False
    for (frameno, frame) in groups:    
        print(frameno)
        bboxval = [0,0,1,1]
        if frameno == len(groups)-1:
            print("last")
            print(len(dfproc.index)%50)
            if len(dfproc.index)%50 <=30:
                printsigs = True
                bboxval = [0,0,1,1]


        # frame.to_csv("%s.csv" % frameno)
        fig, ax =plt.subplots(figsize=(8.5,11))
        # fig.subplots_adjust(top=0.1)
        # ax.suptitle('Kenichi Security (Okada)')
        ax.set_title("Kenichi Security (Okada)\nBi-monthly BillingReport (date1) to (date2)\n",fontsize=10)
        # plt.title('Mean WRFv3.5 LHF\n', fontsize=20)
        
        ax.axis('tight')
        ax.axis('off')
        the_table = ax.table(cellText=frame.values,colLabels=frame.columns,cellLoc='left',loc='center')#bbox=bboxval)#,loc='top')
        the_table.auto_set_font_size(False)
        the_table.set_fontsize(6)
        the_table.auto_set_column_width(col=[0,1,2,3,4])

        cellDict = the_table.get_celld()
        for i in range(0,len(frame.columns)):
            cellDict[(0,i)].set_height(.03)
            for j in range(1,len(frame.index)+1):
                cellDict[(j,i)].set_height(.02)
        # the_table[0, col]._text.set_horizontalalignment('left') 
        # pp.savefig(fig, bbox_inches='tight')#,dpi=150)
        if printsigs:
            # footer = the_table.add_cell(-1, 0, 'Footer Text', loc='center', edgecolor='none', facecolor='none')
            signatoriestext = "________\t\t_________\t\t________\t\t________\njames som\t\tasldkfjh\t\tasldkfjhf\t\tasldkfjhf"
            # plt.figtext(0.5, 0.2, signatoriestext, ha="center", fontsize=18)
            # Set the table title
            # ax.text(0.5, 0.95, 'Table Title', ha='center', va='center', fontsize=14, fontweight='bold')

            # Set the table footer
            ax.text(0.001, -0.06, signatoriestext, ha='center', va='center', fontsize=10)#, color='gray', bbox=dict(boxstyle='square', facecolor='white'))

        
        pp.savefig(fig, bbox_inches='tight')#,dpi=150)
    
    if  True:
        # import matplotlib.pyplot as plt

        # Create a figure and axis
        fig, ax = plt.subplots(figsize=(11,8.5))

        # Set the number of boxes
        num_boxes = 4

        # Set the starting x-coordinate for the first box
        x_start = 0#.001

        # Calculate the width of each box based on the available space
        box_width = (1 - (num_boxes - 1) * 0.1) / num_boxes

        # Set the y-coordinate for the boxes
        y = 0.45
        y1 = 0.6
        y2 = 0.42

        # Set the box height and spacing
        box_height = 0.2
        box_spacing = 0.13

        # Set the box titles
        box_titles = ['Chistianne Kuizon', 'Arenel Abella', 'Roel Lumabao', 'PBGEN Thomas R. Fias jr. (Ret)']
        box_titles1 = ['Prepared by:', 'Checked by:', 'Noted by:', 'Approved by:']
        box_titles2 = ["General Manager","Admin, External Security Opeerations","External Security Operations Manager","Director Physical Security/SSD"]

        # Add the boxes and titles
        for i in range(num_boxes):
            # Calculate the x-coordinate for the current box
            x = x_start + (box_width + box_spacing) * i
            # ax.plot([x_start, x_start+0.1], [y1, y1], color='black')
            
            # Add the box with a bounding box
            ax.text(x, y1, '', bbox=dict(boxstyle='square', facecolor='white', edgecolor='black', linewidth=1))
            
            # Add the title inside the box
            ax.text(x + box_width / 2, y1 + box_height / 2, box_titles1[i]+"      ", ha='right', va='top', fontsize=7)

            ax.text(x, y, '', bbox=dict(boxstyle='square', facecolor='white', edgecolor='black', linewidth=1))
            
            # Add the title inside the box
            ax.text(x + box_width / 2, y + box_height / 2, box_titles[i], ha='center', va='center', fontsize=7)
            # Add the box with a bounding box
            ax.text(x, y2, '', bbox=dict(boxstyle='square', facecolor='white', edgecolor='black', linewidth=1))
            
            # Add the title inside the box
            ax.text(x + box_width / 2, y2 + box_height / 2, box_titles2[i], ha='center', va='center', fontsize=7, weight='bold')
            # ax.plot([x, x + box_width], [y+0.3,y+0.3], color='black')
            ax.axis('off')
        pp.savefig()

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


