import pandas as pd

ref = r'D:\Documents\Python\files\Reference_Compare.xlsx'
opt = input("choose sheet: \n1. SG, 2. Sector, 3. Lvl, 4. Department, 5. Section, 6. Division, 7. Conso JG, 8. Position, 9. Company\nAnswer: ")
sheet = ['SG', 'Sector', 'Lvl', 'Department', 'Section',
         'Division', 'Conso JG', 'Position', 'Company']
index = int(opt)
df = pd.read_excel(
    ref, header=None, sheet_name=sheet[index-1], usecols='K:R').dropna(how='all').fillna('')
print(df)

incorrect_jg1 = ['C1-EQV', 'C2-EQV', 'Eqv.M1-EQV',
                 'Eqv.M2-EQV', 'Est. A1-EQV', 'Est. A3-EQV', 'NT1 /NT5-EQV']
correct_jg1 = ['Est. C1', 'Est. C2', 'P6', 'P7',
               'Est. A1-EQV', 'Est. A3-EQV', 'Est. NT1/NT5']
incorrect_jg2 = ['Est. A1A2-EQV', 'Est. A2-EQV'
                 'Est. C1-EQV', 'Est. C2-EQV', 'Est. D1-EQV', 'Est. D2-EQV'
                 'Est. D3-EQV', 'Est. M1-EQV', 'Est. NE2-EQV', 'Est. NE4-EQV', 'Est. NE6-EQV'
                 ]
incorrect_jg3 = ['Eqv D2-EQV', 'Eqv D3-EQV', 'Eqv M1-EQV', 'Eqv M2-EQV', 'Eqv NT1/NT5-EQV', 'Eqv NT4/NT5-EQV'
                 'Eqv. C2-EQV', 'Eqv. D1-EQV', 'Eqv. D2-EQV', 'Eqv. D3-EQV', 'Eqv. M1-EQV', 'Eqv. M2-EQV'
                 ]

print(df.iloc[:, 0])
# print(df_zp.iloc[:, 4])
# for i, row in df_zp.iloc[:, 4].iteritems():
#     if(row in incorrect_jg1):
#         for j, ic in enumerate(incorrect_jg1):
#             if row == incorrect_jg1[j]:
#                 df_zp.iloc[i, 4] = correct_jg1[j]
#     elif(row in incorrect_jg2):
#         row = 'Est. ' + row[5:-4]
#         df_zp.iloc[i, 4] = row
#     elif row in incorrect_jg3:
#         if row[:4] == 'Eqv.':
#             row = row[:-4]
#             df_zp.iloc[i, 4] = row
#         else:
#             if row == 'Eqv NT1/NT5-EQV' or row == 'Eqv NT4/NT5-EQV':
#                 new = row[:3] + '. ' + row[4:9]
#                 df_zp.iloc[i, 4] = new
#             else:
#                 new = row[:3] + '.' + row[4:6]
#                 df_zp.iloc[i, 4] = new
#     else:
#         new = 'Eqv. ' + row[:-4]
#         df_zp.iloc[i, 4] = new


# df_zp = df[[2, 3]].dropna(how='all')

# df_eq = df.loc[df[2][-3:] == 'EQV']
# print(df_eq)
# print(type(df_zp))
