import pandas as pd


month = input('Select month:\n1. April (4)\n2. May (5)\nAnswer: ')
if month == '4':
    zh_4 = r'D:\Documents\Python\files\April\ZHPLA_april.xlsx'
    zp_4 = r'D:\Documents\Python\files\April\ZPDEV_april.xlsx'
    ta_4 = r'D:\Documents\Python\files\April\TA_april.xlsx'
    df_zh = pd.read_excel(
        zh_4, header=None, sheet_name='ZHPLA_ALL', usecols='A:B')
    df_zp = pd.read_excel(
        zp_4, header=None, sheet_name='Sheet1', usecols='A:B')
    df_ta = pd.read_excel(ta_4, header=None, sheet_name='Sheet1')
else:
    opt = input("insert input: \n1. Full DB\n2. Partial DB\nAnswer: ")
    if opt == '1':
        zhpla_path = r'D:\Documents\Python\files\ZHPLA_MAY20.xlsx'
        zpdev_path = r'D:\Documents\Python\files\ZPDEV_MAY20.xlsx'
        df_zh = pd.read_excel(zhpla_path, header=None,
                              sheet_name='Sheet1', skiprows=3)
        df_zp = pd.read_excel(zpdev_path, header=None,
                              sheet_name='Sheet1', skiprows=3)
        filnam = "TA_input1.csv"
    elif opt == '2':
        zhpla_path = r'D:\Documents\Python\files\zh.xlsx'
        zpdev_path = r'D:\Documents\Python\files\zp.xlsx'
        df_zh = pd.read_excel(zhpla_path, header=None, sheet_name='Sheet1')
        df_zp = pd.read_excel(zpdev_path, header=None, sheet_name='Sheet1')
        filnam = "TA_input2.csv"

# print(df_zh.iloc[:, 0])
print("original number of ZH row: ", len(df_zh.iloc[:, 0]) - 1)
print("original number of ZP row: ", len(df_zp.iloc[:, 0]) - 1)
print("original number of TA row: ", len(df_ta.iloc[:, 0]) - 1)

id_zh = [row for i, row in df_zh.iloc[:, 0].iteritems() if not i ==
         0 and not type(row) == float]
id_zp = [row for i, row in df_zp.iloc[:, 0].iteritems() if not i ==
         0 and not type(row) == float]
id_ta = [row for i, row in df_ta.iloc[:, 0].iteritems() if not i ==
         0 and not type(row) == float]

print("ZH | no of valid id: {}, invalid/nan: {}".format(
    len(id_zh), (len(df_zh.iloc[:, 0]) - 1 - len(id_zh))))
print("ZP | no of valid id: {}, invalid: {}".format(
    len(id_zp), (len(df_zp.iloc[:, 0]) - 1 - len(id_zp))))
# print(id_zh)

zh_special = [staffz for staffz in id_zh if staffz not in id_zp]
zp_special = [staffz for staffz in id_zp if staffz not in id_zh]
print("\nOnly exist in ZH: ", len(zh_special))
print("Only exist in ZP: ", len(zp_special))

allid = df_zh.iloc[1:, 0].to_list()
print('len of allid: ', len(allid))
# allid is equivalent to id_zh
combined_db = id_zh
for idz in id_zp:
    if idz not in combined_db:
        combined_db.append(idz)

print('len of combined db: ', len(combined_db))
print('len of ta: ', len(id_ta))


def getDuplicates(idlist, title):
    seen = set()
    uniq = []
    duplicates = []
    occurence = 0
    for x in idlist:
        if x not in seen:
            uniq.append(x)
            seen.add(x)
        else:
            duplicates.append(x)
    print(title, " | number of duplicates: ", len(duplicates))
    print(duplicates[0:10])


getDuplicates(id_zh, 'ZH')
getDuplicates(id_zp, 'ZP')
getDuplicates(id_ta, 'TA')
