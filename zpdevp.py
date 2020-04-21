import pandas as pd
import os


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
        if(pd.isna(row[0])):
            pass
        else:
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
        if i == 0:
            pass
        else:
            dk = dk.drop([0], axis=0)
        dk = reindex_row(dk)
        dflist[i] = dk
    return dflist


def create_columnName_list(df):
    list_colname = []
    for i, columnName in df.iteritems():
        list_colname.append(columnName[0])
    list_colname.extend(['PPA 5 Years', 'SG (YM)'])
    print(list_colname)
    print("length of colname: %d" % len(list_colname))
    return list_colname


def remove_header(df):
    df = df.drop([0])  # drop row of column name
    df = df.reset_index(drop=True)
    df = df.fillna("")
    return df


def addcolumn_ppa5yrs(df):
    column_42 = []
    ppa5yrs_col = ['40', '38', '36', '34', '32']
    for i, row in df.iterrows():
        text_ppa5yrs = ''
        for year in ppa5yrs_col:
            el = df.iloc[i][year]
            if(el == ''):
                pass
            else:
                el = str(el) + ', '
                text_ppa5yrs += el
        if(text_ppa5yrs[-2:] == ', '):
            text_ppa5yrs = text_ppa5yrs[:-2]
        column_42.append(text_ppa5yrs)
    df.insert(42, '42', column_42, True)
    return df


def addcolumn_SGym(df):
    column_43 = []
    semicolon = ';'
    for i, sg in df.iterrows():
        sg_text = df.iloc[i]['25']
        semicolon_index = sg_text.find(semicolon)
        sg_date = sg_text[:semicolon_index-1]
        sg_ym = sg_text[semicolon_index+1:]
        df.at[i, '25'] = sg_date
        column_43.append(sg_ym)
    df.insert(43, '43', column_43, True)
    return df


def arrange_renameCol(df, list_columnname):
    index_order = [0, 1, 3, 16, 17, 18, 32, 34, 36, 38, 40, 42, 4, 5, 6, 7, 8,
                   9, 10, 11, 19, 24, 23, 22, 21, 20, 25, 43, 29, 15, 30, 31, 26, 28, 27]
    column_order = []
    column_name = []
    for i in index_order:
        s = str(i)
        column_order.append(s)  # basically same as index order but as string
        column_name.append(list_columnname[i])
    column_date = [4, 25, 29, 15, 30, 31]
    for col in column_date:
        c = str(col)
        for i, row in df.iterrows():
            date = df.iloc[i][c]
            date = date.replace('.', '/')
            df.at[i, c] = date
    # part 1
    df = df.reindex(columns=column_order)
    df = df.sort_values(by=['0'], ascending=True)  # takde pun takpe
    # part 2
    df.columns = column_name
    df.rename(columns={'Personnel Number': 'Staff Number',
                       'Formatted Name of Employee or Applicant': 'Staff Name',
                       'Gender Key Desc': 'Gender', 'Sal. Grade': 'SG', 'Job Grade': 'JG',
                       'Date Post': 'Date in Position', 'Join Dept': 'Date in Department',
                       'Join Div.': 'Date in Division', 'Join Comp.': 'Date in Company'}, inplace=True)
    return df


mainpath = os.getcwd()
# zpdev_path = os.path.join(mainpath, "database/input/", "zpdevc.xlsx")
input_path = r'D:\Documents\Python\project-2\database\input\zpdevc.xlsx'
output_path = os.path.join(mainpath, "database/output/", "cleanzpdev.xlsx")
df = pd.read_excel(input_path, skiprows=4, nrows=None)

df = reindex_column(df)
dflist = slicebigdf(df)
dflist = adjustalldf(dflist)
bigdf = pd.concat(dflist).reset_index(drop=True)
bigdf.to_csv("zpdev_before.csv", index=None)
list_colname = create_columnName_list(bigdf)
bigdf = remove_header(bigdf)
bigdf = addcolumn_ppa5yrs(bigdf)
bigdf = addcolumn_SGym(bigdf)
bigdf = arrange_renameCol(bigdf, list_colname)

print(bigdf)
print("Writing to excel...")
# bigdf.to_excel(output_path, index=None, header=None)
bigdf.to_csv("zpdev_after.csv", index=None)
print("DONE")
