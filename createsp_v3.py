import pandas as pd
import os
import math
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt


def reindex_column(df):
    colname = list(df.columns)
    colname = [i.replace(i, str(colname.index(i))) for i in colname]
    # colname = [i.replace(i[-2:],'') for i in colname] # replace ".1" to ""
    df.columns = colname
    return df


def df_tolist(df):
    sheet = []
    for column in df:
        cc = df[column]
        list_ = [column]  # with header
        list_ += cc.values.tolist()
        sheet.append(list_)
        # print(li_st)
        # print('------------------')
    return sheet


def getFileInfo(filepath):
    file = pd.ExcelFile(filepath)
    filename = Path(filepath).stem
    sheetname_list = file.sheet_names
    return filename, sheetname_list


def splitBelow9Rows(column, columnlength):
    breakindex = math.ceil(columnlength/2)
    df_a = column.iloc[:breakindex].reset_index(drop=True)
    df_b = column.iloc[breakindex:].reset_index(drop=True)
    return df_a, df_b


def splitBelow15Rows(column, columnlength):
    breakindex = math.ceil(columnlength/3)
    df_a = column.iloc[:breakindex]
    df_a2 = column.iloc[breakindex:]
    df_b, df_c = splitBelow9Rows(df_a2, len(df_a2))
    return df_a, df_b, df_c


directory = os.getcwd()
filepath = r'D:\Documents\Python\project-2\database\input\SP_GHRM.xlsx'
filename, sheetname_list = getFileInfo(filepath)
# sheetcount = len(sheetname_list)

data_df_list = []
id_df_list = []
for i, sheetname in enumerate(sheetname_list):
    sheetdf = pd.read_excel(filepath, sheet_name=sheetname,
                            skiprows=2, nrows=None)
    data_df = sheetdf.iloc[:, [6, 8, 10]].dropna(how='all')
    id_df = sheetdf.iloc[:, [0, 1, 2]].dropna(how='all')
    data_df_list.append(data_df)
    id_df_list.append(id_df)

# prs = Presentation()
# title_slide_layout = prs.slide_layouts[0]
# title_content_layout = prs.slide_layouts[1]

for i, data_df in enumerate(data_df_list):
    data_df = reindex_column(data_df)
    columnlenlist = []
    sheet = df_tolist(data_df)
    rowcount, columncount = data_df.shape
    print("CASE %d" % i)
    print(data_df)
    collen_srs = data_df.count()  # contains length of each column
    collen_srs.reset_index(drop=True)
    print("columnlengthsrs: \n{}".format(collen_srs))
    idx_max = int(collen_srs.idxmax())  # which line is the longest
    print("idx max = ", idx_max)
    # longestcolumn = collen_srs[idx_max]  # how long is the line

    for j in collen_srs:
        # later we want to test if the list length is still 3 or not
        # going to act as indicator
        columnlenlist.append(j)

    if(idx_max == 1):
        d1, d2, d3 = columnlenlist
        df1 = data_df.iloc[:, 0].dropna(how='all')
        df2 = data_df.iloc[:, 1].dropna(how='all')
        df3 = data_df.iloc[:, 2].dropna(how='all')
        priority_dict = {0: df1, 1: df2, 2: df3}
        print("indexes: {}, {}, {}".format(1, 2, 3))
        for pd in priority_dict:
            dfpd = priority_dict[pd]
            collen = len(dfpd)
            if collen > 5:
                if collen <= 9:
                    dfa, dfb = splitBelow9Rows(dfpd, collen)
                    splitlist = [dfa, dfb]
                    priority_dict[pd] = splitlist
                else:
                    if collen <= 15:
                        dfa, dfb, dfc = splitBelow15Rows(dfpd, collen)
                        splitlist = [dfa, dfb, dfc]
                        priority_dict[pd] = splitlist
                    else:
                        # data is bigger than 15, need to split to 2
                        dfpage1 = dfpd.iloc[:15]
                        dfpage2 = dfpd.iloc[15:]
                        pass
            else:
                priority_dict[pd] = [dfpd]
        # print(priority_dict)

        # repeat the same as idxmax 0 or 2
    else:
        if(columnlenlist[0] == columnlenlist[2]):
            indexfirst, indexsecond, indexthird = 1, 2, 0
        elif(columnlenlist[0] == columnlenlist[1]):
            indexfirst, indexsecond, indexthird = 2, 1, 0
        else:
            if(idx_max == 0):  # 0
                # if(columnlenlist[0] == columnlenlist[2]):
                #     d1, d2 = columnlenlist[1], columnlenlist[2]
                # else:
                d1, d2 = columnlenlist[1], columnlenlist[2]
            elif(idx_max == 2):
                d1, d2 = columnlenlist[0], columnlenlist[1]
            dofirst = min(d1, d2)  # length of shortest column
            dosecond = max(d1, d2)  # length of mid column
            dothird = columnlenlist[idx_max]

            indexfirst = columnlenlist.index(dofirst)  # index of shortest
            indexsecond = columnlenlist.index(dosecond)  # index of mid
            indexthird = columnlenlist.index(dothird)  # index of longest
            if(indexsecond == indexfirst):
                for i, cl in enumerate(columnlenlist):
                    if not(i == idx_max or i == columnlenlist.index(dofirst)):
                        indexsecond = i

        print("indexes: {}, {}, {}".format(
            indexfirst, indexsecond, indexthird))

        hold_df = data_df.iloc[:, indexthird].dropna(how='all')
        df1 = data_df.iloc[:, indexfirst].dropna(how='all')
        df2 = data_df.iloc[:, indexsecond].dropna(how='all')

        priority_dict = {indexfirst: df1,
                         indexsecond: df2,
                         indexthird: hold_df}

        for pd in priority_dict:
            print("pd: {}".format(pd))
            collen = len(priority_dict[pd])
            dfpd = priority_dict[pd]
            if collen > 5:
                # print("panjang")
                # process split the column into 2 rows here
                if collen <= 9:
                    dfa, dfb = splitBelow9Rows(dfpd, collen)
                    splitlist = [dfa, dfb]
                    priority_dict[pd] = splitlist
                else:
                    if collen <= 15:
                        dfa, dfb, dfc = splitBelow15Rows(dfpd, collen)
                        splitlist = [dfa, dfb, dfc]
                        priority_dict[pd] = splitlist
                    else:
                        # do 2 page, split the first 15 first
                        dfpage1 = dfpd.iloc[:15]
                        dfpage2 = dfpd.iloc[15:]
                        print(dfa, "\n", dfb)
                        # dfa do the collen <= 15
                        # dfb do the collen <= 9
            else:
                priority_dict[pd] = [dfpd]
    print(priority_dict)
