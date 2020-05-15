import pandas as pd
import os
import datetime


def reindex_column(df):
    colname = list(df.columns)
    colname = [i.replace(i, str(colname.index(i))) for i in colname]
    # colname = [i.replace(i[-2:],'') for i in colname] # replace ".1" to ""
    df.columns = colname
    return df


def reindex_row(df):
    rowlength = df.shape[0]
    newindex = pd.Series([i for i in range(rowlength)])
    df = df.set_index(newindex)
    return df


def getindexlist(df):
    # list contain index of new group of data, detects 'zhpla report' from column 0/A
    listofindex = []
    for i, row in df.iterrows():
        if not pd.isna(row[0]):
            if(row[0] == 'HR Flexible Report'):
                listofindex.append(i)
    # print("length of listofi: {}".format(len(listofindex)))
    return listofindex, len(listofindex)


def risecolumn(dx):
    dx = dx.drop([0], axis=0)
    dxlen = len(dx)
    newindex = [i for i in range(dxlen)]
    dx.index = newindex
    for i, v in dx.iteritems():
        if(i == 0):
            pass
        elif(i == 1 or i == 2):
            v = float('nan')
            dx[i] = v
        else:
            if(i % 2 == 0):
                v = float('nan')
                dx[i] = v
            else:  # odd is posid
                pass
    return dx


def separate_posid(dx):
    for i, v in dx.iteritems():
        if(i == 0):
            pass
        elif(i == 1 or i == 2):
            v = float('nan')
            dx[i] = v
        else:
            if(i % 2 == 0):
                v = float('nan')
                dx[i] = v
            else:  # odd is posid
                pass
    return dx


def slicebigdf(df):
    listofindex, count = getindexlist(df)
    dflist = []
    for i in range(count):
        st = listofindex[i]  # st is the starting index of the group
        if i < count-1:  # if st is not the last element in the list, then we can set the endpoint
            en = listofindex[i+1]  # endpoint is the next element(id)
            df_ = df.iloc[st+3:en]  # slice the big df into the specified range
        else:  # if st is the last element
            df_ = df.iloc[st+3:]  # no need to set the endpoint
        df_ = df_.drop(['0', '2'], axis=1)  # drop the empty column at index 4
        df_ = df_.dropna(how='all')
        df_ = reindex_row(df_)
        df_ = reindex_column(df_)
        dflist.append(df_)
    print("length of dflist(number of group): {}".format(len(dflist)))
    return dflist


def adjustalldf(dflist):
    print("Adjusting columns...")
    for i, dk in enumerate(dflist):
        if not i == 0:
            dk = dk.drop([0], axis=0)
        dk = reindex_row(dk)
        dflist[i] = dk
    return dflist


def create_columnName_list(df):
    list_colname = [columnName[0] for i, columnName in df.iteritems()]
    return list_colname


def set_header(df, columnName):
    df = df.drop([0]).reset_index(drop=True)  # drop row of column name
    df.columns = columnName
    return df


def addcolumn_ppa5yrs(df):
    column_ppa = []
    ppa5yrs_col = ['PPA-15', 'PPA-16', 'PPA-17', 'PPA-18', 'PPA-19']
    for i, row in df.iterrows():
        text_ppa5yrs = ''
        for i, year in enumerate(ppa5yrs_col):
            el = df.loc[i, year]
            if not el == '' or type(el) is float:
                if not i == len(ppa5yrs_col)-1:
                    el = str(el) + ', '
                text_ppa5yrs += str(el)
        if(text_ppa5yrs[-2:] == ', '):
            text_ppa5yrs = text_ppa5yrs[:-2]
        column_ppa.append(text_ppa5yrs)
    df['PPA'] = column_ppa
    return df


def addcolumn_SGym(df):
    column_sgym = []
    for i, column in df.iteritems():
        if(i == 'SG Since'):
            for j, row in column.iteritems():
                separator = row.find(';')
                sg_date = row[:separator-1]
                sg_ym = row[separator+1:]
                df.loc[j, 'SG Since'] = sg_date
                column_sgym.append(sg_ym)
    df['SG (YM)'] = column_sgym
    return df


def addcolumn_age(df):
    column_date = ['Date of Birth', 'Date in Department',
                   'SG Since', 'Date in Position', 'Date in Division', 'Date in Company']
    for colname, column in df.iteritems():
        if colname in column_date:
            for i, row in column.iteritems():
                if type(row) is str:
                    row = row.replace('.', '/')
                    df.loc[i, colname] = row.replace('.', '/')
    column_age = []
    format_str = '%d/%m/%Y'  # The format
    today = datetime.date.today()
    for i, value in df['Date of Birth'].iteritems():
        dob = datetime.datetime.strptime(value, format_str)
        age = today.year - dob.year - \
            ((today.month, today.day) < (dob.month, dob.day))
        column_age.append(age)
    df['Age'] = column_age


def arrange_renameCol(df):

    index_order = [0, 1, 3, 16, 17, 18, 32, 34, 36, 38, 40, 42, 4, 5, 6, 7, 8,
                   9, 10, 11, 19, 24, 23, 22, 21, 20, 25, 43, 29, 15, 30, 31, 26, 28, 27]
    # for i, colname in enumerate(df.columns):
    #     print(colname)

    # print(df['Date of Birth'])
    df = addcolumn_age(df)

    # part 1
    # df = df.reindex(columns=column_order)
    # df = df.sort_values(by=['0'], ascending=True)  # takde pun takpe
    # part 2
    # df.columns = column_name
    return df


def set_input_output(opt):
    input_path = r'D:\Documents\Python\project-2\database\input\zpdev-may2020.xlsx' if opt == '1' else r'D:\Documents\Python\project-2\database\input\zpdevc.xlsx'
    output_path = r'D:\Documents\Python\project-2\database\output\zp_full.xlsx' if opt == '1' else r'D:\Documents\Python\project-2\database\output\zp_partial.xlsx'
    return input_path, output_path


def clean_rawdf(df):
    df = reindex_column(df)
    dflist = slicebigdf(df)
    dflist = adjustalldf(dflist)
    cleandf = pd.concat(dflist).reset_index(drop=True)
    return cleandf


def format_zpdev(df):
    # list_colname = create_columnName_list(df)
    list_colname = [columnName[0] for i, columnName in df.iteritems()]
    df = set_header(df, list_colname)
    df = df.rename(columns={'Personnel Number': 'Staff Number',
                            'Formatted Name of Employee or Applicant': 'Staff Name',
                            'Gender Key Desc': 'Gender', 'Sal. Grade': 'SG', 'Job Grade': 'JG',
                            'Date Post': 'Date in Position', 'Join Dept': 'Date in Department',
                            'Join Div.': 'Date in Division', 'Join Comp.': 'Date in Company'}, inplace=True)
    df = addcolumn_ppa5yrs(df)
    df = addcolumn_SGym(df)
    df = addcolumn_age(df)
    df = arrange_renameCol(df)
    return df


mainpath = os.getcwd()
opt = input('1. Full ZP\n2. Partial ZP\nAnswer: ')
input_path, output_path = set_input_output(opt)


df = pd.read_excel(input_path, skiprows=4, nrows=None)
cleandf = clean_rawdf(df)

finaldf = format_zpdev(cleandf)

# print(finaldf)
print("Writing to excel...")
# cleandf.to_excel(output_path, index=None)
# df.to_csv("zpdev_after.csv", index=None)
print("DONE")
