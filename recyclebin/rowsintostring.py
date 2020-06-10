import pandas as pd
# replace path/name here
filepathname = r'D:\Documents\Python\project-2\database\input\staff_id(1).xlsx'
file = pd.ExcelFile(filepathname)
df = pd.read_excel(file, sheet_name='Sheet1', usecols='A')
df = df.applymap(str)
print(df)
# s = df['ID'].str.cat(sep=",")
s = df.iloc[:, 0].str.cat(sep=",")  # get the first column
# print(s)
id_textfile = open("staff_id_text(1).txt", "w")
id_textfile.write(s)
id_textfile.close()


def IDlistToString(filepath, df):
    pass
