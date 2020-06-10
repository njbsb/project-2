import pandas as pd
import cleanzp_lib as cleanzp

zp_junepath = r'D:\Documents\Python\project-2\database\input\raw_zpdev_june2020_v2.xlsx'
outputpath = r'D:\Documents\Python\project-2\database\output\new_zpdev_june2020_v2.xlsx'

df_data = pd.read_excel(zp_junepath, sheet_name='Data')
df_id = pd.read_excel(zp_junepath, sheet_name='IDs')

mergedf = pd.merge(df_id, df_data, on=[
                   'Personnel Number', 'Formatted Name of Employee or Applicant'], how='outer')
mergedf = mergedf.sort_values(by='Personnel Number')
df = cleanzp.addcolumn_ppa5yrs(mergedf)
df = cleanzp.addcolumn_SGym(df)
df = cleanzp.addcolumn_age(df)
df = cleanzp.arrange_renameCol(df)
print(df)
df.to_excel(outputpath, index=None)
