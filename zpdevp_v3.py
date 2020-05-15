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


def set_header(df, colname):
    df = df.drop([0]).reset_index(drop=True)  # drop row of column name
    df.columns = colname
    return df


def addcolumn_ppa5yrs(df):
    column_ppa = []
    ppa5yrs_col = ['PPA-15', 'PPA-16', 'PPA-17', 'PPA-18', 'PPA-19']
    for i, row in df.iterrows():
        text_ppa5yrs = ''
        for j, year in enumerate(ppa5yrs_col):
            ppa = df.loc[i, year]
            if not ppa == '' or type(ppa) is float:
                if not j == len(ppa5yrs_col) - 1:
                    ppa = str(ppa) + ', '
                text_ppa5yrs += str(ppa)
        column_ppa.append(text_ppa5yrs)
    df['PPA'] = column_ppa
    return df


def addcolumn_SGym(df):
    column_sgym = []
    for i, column in df.iteritems():
        if i == 'SG Since':
            for j, row in column.iteritems():
                separator = row.find(';')
                sg_date = row[:separator-1]
                sg_ym = row[separator+1:]
                df.loc[j, 'SG Since'] = sg_date
                column_sgym.append(sg_ym)
    df['SG (YM)'] = column_sgym
    return df


def arrange_renameCol(df):
    column_date = ['Date of Birth', 'Join Dept',
                   'SG Since', 'Date Post', 'Join Div.', 'Join Comp.']
    for colname, column in df.iteritems():
        if colname in column_date:
            for i, row in column.iteritems():
                if not type(row) is float:
                    row = row.replace('.', '/')
                    df.loc[i, colname] = row
    df.rename(columns={'Personnel Number': 'Staff Number',
                       'Formatted Name of Employee or Applicant': 'Staff Name',
                       'Gender Key Desc': 'Gender', 'Sal. Grade': 'SG', 'Job Grade': 'JG',
                       'Date Post': 'Date in Position', 'Join Dept': 'Date in Department',
                       'Join Div.': 'Date in Division', 'Join Comp.': 'Date in Company', 'Join Date': 'Date in PETRONAS'}, inplace=True)
    column_order = ['Staff Number', 'Staff Name', 'Gender', 'Age', 'SG', 'Lvl', 'JG', 'PPA-19', 'PPA-18', 'PPA-17', 'PPA-16', 'PPA-15', 'PPA',
                    'Date of Birth', 'Country of Birth', 'Country of Birth Desc', 'State', 'State Desc', 'Birthplace', 'Nationality', 'Nationality Desc', 'Position',
                    'Section', 'Department', 'Division', 'Sector', 'Company', 'SG Since', 'SG (YM)', 'Date in PETRONAS', 'Date in Position', 'Date in Department', 'Date in Division',
                    'Date in Company', 'Pos. SKG', 'Home SKG', 'Work Center', 'Religious Denomination Key', 'Religious Denomination Key Desc']
    df = df[column_order]
    return df


def addcolumn_age(df):
    today = datetime.date.today()
    format_str = '%d/%m/%Y'
    column_age = []
    for i, row in df['Date of Birth'].iteritems():
        if not row == float('nan'):
            row = row.replace('.', '/')
            dob = datetime.datetime.strptime(row, format_str)
            age = today.year - dob.year - \
                ((today.month, today.day) < (dob.month, dob.day))
        else:
            age = 0
        column_age.append(age)
    df['Age'] = column_age
    return df


def clean_rawdf(df):
    df = reindex_column(df)
    dflist = slicebigdf(df)
    dflist = adjustalldf(dflist)
    cleandf = pd.concat(dflist).reset_index(drop=True)
    # bigdf.to_csv("zpdev_before.csv", index=None)
    return cleandf


def format_zpdev(bigdf):
    list_colname = [colname[0] for i, colname in bigdf.iteritems()]
    bigdf = set_header(bigdf, list_colname)  # done
    bigdf = addcolumn_ppa5yrs(bigdf)  # done
    bigdf = addcolumn_SGym(bigdf)
    bigdf = addcolumn_age(bigdf)
    bigdf = arrange_renameCol(bigdf)
    return bigdf


def set_input_output(opt):
    input_path = r'D:\Documents\Python\project-2\database\input\zpdev-may2020.xlsx' if opt == '1' else r'D:\Documents\Python\project-2\database\input\zpdevc.xlsx'
    output_path = r'D:\Documents\Python\project-2\database\output\zp_full.xlsx' if opt == '1' else r'D:\Documents\Python\project-2\database\output\zp_partial.xlsx'
    return input_path, output_path


mainpath = os.getcwd()
opt = input('1. Full ZP\n2. Partial ZP\nAnswer: ')

input_path, output_path = set_input_output(opt)

df = pd.read_excel(input_path, skiprows=4, nrows=None)

cleandf = clean_rawdf(df)

finaldf = format_zpdev(cleandf)

print(finaldf)
print("Writing to excel...")
finaldf.to_excel(output_path, index=None, header=None)
# bigdf.to_csv("zpdev_after.csv", index=None)
print("DONE")
