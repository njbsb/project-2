import pandas as pd
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches, Pt
# from pptx_sp import add_slide, create_table, create_table2, iter_cells, resize_table_font, normalize_df, reindex_column
import pptx_sp as ps

print("-------------------------------------------------------------------------------- NEW COMMAND ------------------------------------------------------------------")

sp_input_path = r'C:\Users\mnajib.suib\Documents\Python\project-2\database\input\sp-ptc1-2020.xlsx'
sp_inputfile = pd.ExcelFile(sp_input_path)
sp_numberofsheet = len(sp_inputfile.sheet_names)
print("Number of sheets: %d" %sp_numberofsheet)

df_list = []
sheetname_ = []
for s in range(sp_numberofsheet):
    df_spexcel = pd.read_excel(sp_input_path, sheet_name=s, skiprows=2, usecols='F:K', nrows=None)
    df_spexcel = df_spexcel.iloc[:,[1,3,5]]
    df_list.append(df_spexcel)

with pd.ExcelWriter('spexcel_output1.xlsx') as writer: #pylint: disable=abstract-class-instantiated
    for n, df in enumerate(df_list):
        df.to_excel(writer, 'Sheet%s' % (n+1), index=False)
    writer.save()
# sp excel file is created, df_list contain list of dataframe

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
title_content_layout = prs.slide_layouts[1]
sl_title, sh_title = ps.add_slide(prs, title_slide_layout, "Title")

new_dflist = []
for df in df_list:
    df = ps.reindex_column(df)
    print("----------------------------------------------------------------------------------\nNEW SHEET")
    print("Sheet/dataframe shape: {}, rowcount: {}".format(str(df.shape), df.shape[0])) # tuple of (rowlength, collength)

    newlist2 = []
    df, newlist2 = ps.normalize_df(df)
    slide, shape = ps.add_slide(prs, title_content_layout, "HOD .....")
    ps.create_table(slide, newlist2, 1, 10)

sl_end, sh_end = ps.add_slide(prs, title_slide_layout, "End")

prs.save('createsp.pptx')
