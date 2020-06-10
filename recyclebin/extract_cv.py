import os
import PyPDF2 as pypdf
import tabula
import pandas as pd
import camelot

cv_path = r'D:\Documents\Python\project-2\database\output\shazlinacv.pdf'
cv_obj = pypdf.PdfFileReader(cv_path)

# for page in cv_obj.pages:
#     pagetext = page.extractText()
#     print(pagetext + "\n")
df = tabula.read_pdf(cv_path, pandas_options={
                     'header': None})[0]

# dff = df[0]
# df.to_excel("adleencv.xlsx", index=None, header=None)
print(df)
print(type(df))

