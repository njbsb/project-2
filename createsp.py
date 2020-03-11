import pandas as pd
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx_sp import add_slide, create_table, create_table2, iter_cells, resize_table_font

sp_input_path = r'C:\Users\mnajib.suib\Documents\Python\project-2\database\input\sp-ptc1-cs.xlsx'
sp_inputfile = pd.ExcelFile(sp_input_path)
sp_numberofsheet = len(sp_inputfile.sheet_names)
print("Number of sheets: %d" %sp_numberofsheet)

df_list = []
for s in range(sp_numberofsheet):
    df_spexcel = pd.read_excel(sp_input_path, sheet_name=s, skiprows=2, usecols='F:K', nrows=None)
    df_spexcel = df_spexcel.iloc[:,[1,3,5]]
    # df_spexcel = df_spexcel.fillna("")
    # print(df_spexcel)

    df_list.append(df_spexcel)

with pd.ExcelWriter('spexcel_output.xlsx') as writer: #pylint: disable=abstract-class-instantiated
    for n, df in enumerate(df_list):
        df.to_excel(writer, 'Sheet%s' % (n+1), index=False)
    writer.save()

# sp excel file is created, df_list contain list of dataframe

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
title_content_layout = prs.slide_layouts[1]

sl_title, sh_title = add_slide(prs, title_slide_layout, "Title")

for df in df_list:
    sheet_list = []
    sheet_list2 = []

    colname = list(df.columns)
    colname = [i.replace(i[-2:],'') for i in colname] # replace ".1" to ""
    df.columns = colname # rename the column of df

    print("\nSheet/dataframe shape: " + str(df.shape))
    
    for i, (columnName, columnData) in enumerate(df.iteritems()):
        li_st = columnData.values.tolist() # convert column to columnlist
        li_st.insert(0, colname[i]) # insert column name to columnlist
        columnName
        sheet_list.append(li_st) # insert to the sheet list
    
    for i, column in enumerate(df):
        li_st = []
        cc = df[column]
        li_st.append(column)
        li_st = li_st + cc.values.tolist()
        sheet_list2.append(li_st)

    print(df)
    j = df.count()
    jmax = j.idxmax()
    print("column with most row: {}".format(jmax))
    
    
    slide, shape = add_slide(prs, title_content_layout, "HOD .....")
    # create_table(slide, sheet_list, 1, 10)
    # create_table2(slide, df, 1, 10)

sl_end, sh_end = add_slide(prs, title_slide_layout, "End")

prs.save('createsp.pptx')
