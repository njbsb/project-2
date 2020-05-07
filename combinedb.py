import pandas as pd
import os
from datetime import date, datetime
from dateutil.relativedelta import relativedelta


def date_to_duration(date):
    dip = date.date()
    rdelta = relativedelta(date.today(), dip)
    dateYM = str(rdelta.years) + "y" + str(rdelta.months) + "m"
    return dateYM


def drop_unusedColumn(df, kind):
    droplist = []
    if kind == "zh":
        keeplist = [0, 1, 29, 30, 2, 3, 4, 8, 10, 11, 12, 13, 14, 15, 16, 17, 18,
                    19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 32, 33, 34, 35, 37, 38, 43]
    else:
        keeplist = [0, 3, 27, 28, 40, 41, 7, 8, 9, 10, 11, 12, 29,
                    30, 31, 32, 33, 13, 14, 15, 16, 17, 18, 19, 20, 34, 35, 36, 37, 38, 39]
    for i in range(df_zp.shape[1]):
        if i not in keeplist:
            droplist.append(i)
    df = df.drop(columns=droplist)
    return df


def reindex_column(df):
    colname = list(df.columns)
    for i in range(len(colname)):
        colname[i] = i
    df.columns = colname
    # print(df.head())
    return df


def rearrangeTA_column(df_zh, df_zp):
    print("index", df_zh.index)
    zh_col = []
    zp_col = []
    for i, ec in df_zh.iteritems():
        zh_col.append(ec[0])
    for i, ec in df_zp.iteritems():
        zp_col.append(ec[0])
    df_zh = df_zh.drop([0]).reset_index(drop=True)
    df_zp = df_zp.drop([0]).reset_index(drop=True)
    # print(zp_col)
    df_zh.columns = zh_col
    df_zp.columns = zp_col
    df_zp = df_zp.rename(columns={"Personnel Number": "Staff No",
                                  "Pos. SKG": "Pos. SKG Zpdev", "Home SKG": "Home SKG Zpdev"})
    df = pd.merge(df_zh, df_zp, on='Staff No', how='left')
    ori_colname = list(df.columns)
    df = reindex_column(df)  # convert column name to index
    final = [0, 1, 33, 25, 26, 35, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 51, 63, 64, 36, 37, 38, 39, 40, 41, 52, 53, 54, 55, 56, 42, 43, 44,
             45, 46, 47, 48, 49, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 27, 28, 57, 58, 29, 30, 31, 32, 59, 34, 60, 61, 62]
    arr_column = []
    if(len(ori_colname) == len(final)):
        for i in final:
            arr_column.append(ori_colname[i])
    df = df.reindex(columns=final)
    df.columns = arr_column
    return df


directory = os.getcwd()
opt = input("insert input:")
if opt == '1':
    zhpla_path = r'D:\Documents\Python\files\ZHPLA_MAY20.xlsx'
    zpdev_path = r'D:\Documents\Python\files\ZPDEV_MAY20.xlsx'
elif opt == '2':
    zhpla_path = r'D:\Documents\Python\files\zh.xlsx'
    zpdev_path = r'D:\Documents\Python\files\zp.xlsx'
df_zh = pd.read_excel(zhpla_path, header=None, sheet_name='Sheet1')
df_zp = pd.read_excel(zpdev_path, header=None, sheet_name='Sheet1')

df_zp[40] = df_zp[29]
df_zp[41] = df_zp[30]
for i, column in df_zp.iteritems():
    if i == 40 or i == 41:
        for index, row in column.iteritems():
            if not index == 0:
                column[index] = date_to_duration(row)
            else:
                if(i == 40):
                    column[index] = "YiPetronas"
                else:
                    column[index] = "YiPosition"

df_zh = drop_unusedColumn(df_zh, 'zh')
df_zp = drop_unusedColumn(df_zp, 'zp')

df_merge = rearrangeTA_column(df_zh, df_zp)
print(df_merge)
df_merge.to_csv("TAFULL.csv", header=True, index=False)
