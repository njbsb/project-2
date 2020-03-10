import pandas as pd
import os
import xlrd
import openpyxl as op
from openpyxl.utils import get_column_letter, column_index_from_string
import excelconvert as cv
import string

# convert to xlsx

df_path = r'C:\Users\mnajib.suib\Documents\Python\project-2\database\input\zhpla-march2020.xlsx'
df_paths = r'C:\Users\mnajib.suib\Documents\Python\project-2\database\input\1000records.xlsx'
df_ = pd.read_excel(df_path)[4:]
df_ = df_.reset_index(drop=True)
print(df_.shape)
# print(df_)
# maxrow = sheet.max_row
# maxcol = sheet.max_column

def colToExcel(col): # col is 1 based
    excelCol = str()
    div = col 
    while div:
        (div, mod) = divmod(div-1, 26) # will return (x, 0 .. 25)
        excelCol = chr(mod + 65) + excelCol

    return excelCol

print(df_.columns)
for i in range(len(df_)):
    df_.rename(columns={df_.columns[i]: colToExcel(i)}, inplace=True)

print(df_.columns)