import pandas as pd
filepathname = 'cv.xlsx'  # replace path/name here
file = pd.ExcelFile(filepathname)
df = pd.read_excel(file, sheet_name='Sheet2', usecols='A')
df = df.applymap(str)
print(df)
s = df['ID'].str.cat(sep=",")
print(type(df['ID']))
id_textfile = open("id_text.txt", "w")
id_textfile.write(s)
id_textfile.close()
