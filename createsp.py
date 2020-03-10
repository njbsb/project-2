import pandas as pd
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx_sp import add_slide, create_table, iter_cells, resize_table_font

sp_input_path = r'C:\Users\mnajib.suib\Documents\Python\project-2\database\input\sp-ptc1-cs.xlsx'
sp_inputfile = pd.ExcelFile(sp_input_path)
sp_numberofsheet = len(sp_inputfile.sheet_names)
df_list = []
for s in range(sp_numberofsheet):
    print("Sheet: " + str(s))
    df_spexcel = pd.read_excel(sp_input_path, sheet_name=s, skiprows=2, usecols='F:K', nrows=None)
    df_spexcel = df_spexcel.iloc[:,[1,3,5]]
    df_spexcel = df_spexcel.fillna("")
    df_list.append(df_spexcel)
    # print(df_spexcel)

with pd.ExcelWriter('spexcel_output.xlsx') as writer: #pylint: disable=abstract-class-instantiated
    for n, df in enumerate(df_list):
        df.to_excel(writer, 'Sheet%s' % (n+1))
    writer.save()


# sp excel file is created, df_list contain list of dataframe

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
title_content_layout = prs.slide_layouts[1]

sl_title, sh_title = add_slide(prs, title_slide_layout, "Title")

for df in df_list:
    sheet_list = []
    for columnName, columnData in df.iteritems():
        li_st = columnData.values.tolist()
        if(columnName[-2:-1] == "."):
            columnName = columnName[:-2]  
        li_st.insert(0, columnName)
        sheet_list.append(li_st)
    print(sheet_list)
    slide, shape = add_slide(prs, title_content_layout, "HOD .....")
    create_table(slide, sheet_list, 1, 10)

sl_end, sh_end = add_slide(prs, title_slide_layout, "End")

prs.save('createsp.pptx')
