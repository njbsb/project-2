import pandas as pd
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches, Pt
import pptx_sp as ps

print("------------------------------------------------------------------ NEW COMMAND ------------------------------------------------------------------")

sppath = r'C:\Users\mnajib.suib\Documents\Python\project-2\database\input\sp-ptc1-2020.xlsx'
spfile = pd.ExcelFile(sppath)
sheetname = spfile.sheet_names
sheetcount = len(spfile.sheet_names)
print("Number of sheets: %d" % sheetcount)
df_list = []
for s in range(sheetcount):
    df_ = pd.read_excel(sppath, sheet_name=s, skiprows=2,
                        usecols='F:K', nrows=None)
    df_ = df_.iloc[:, [1, 3, 5]]
    df_list.append(df_)
ps.write_excel(df_list, 'sp_output1.xlsx', sheetname)
prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
title_content_layout = prs.slide_layouts[1]
sl_title, sh_title = ps.add_slide(prs, title_slide_layout, "Title")

for i, df in enumerate(df_list):
    df = ps.reindex_column(df)
    print("-----------------------------------------------------\n{}\nSheet/df shape: {}, rowcount: {}".format(
        sheetname[i], str(df.shape), df.shape[0]))
    sheet = []
    df, sheet = ps.normalize_df(df)
    slide, shape = ps.add_slide(prs, title_content_layout, sheetname[i])
    ps.create_table(slide, sheet, 1, 10)
    # ps.create_table_df(slide, df, 1, 10)
    df_list[i] = df

ps.write_excel(df_list, 'sp_output2.xlsx', sheetname)
sl_end, sh_end = ps.add_slide(prs, title_slide_layout, "End")
prs.save('createsp.pptx')
