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
            listofindex.append(i)
    # print("length of listofi: {}".format(len(listofindex)))
    return listofindex, len(listofindex)


def risecolumn(dx):
    # function: to raise sunken column
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
    # function: separate posID and EndDate that are in the same column
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
    # function: identify group of data in original zhpla
    # input: dirty dataframe
    # output: list containing small group of df
    # returns list that contains index to use for slicing, and its length
    listofindex, count = getindexlist(df)
    dflist = []
    for i in range(count):
        st = listofindex[i]  # st is the starting index of the group
        if i < count-1:  # if st is not the last element in the list, then we can set the endpoint
            en = listofindex[i+1]  # endpoint is the next element(id)
            df_ = df.iloc[st+3:en]  # slice the big df into the specified range
        else:  # if st is the last element
            df_ = df.iloc[st+3:]  # no need to set the endpoint
        df_ = df_.drop(['4'], axis=1)  # drop the empty column at index 4
        df_ = reindex_row(df_)
        df_ = reindex_column(df_)
        dflist.append(df_)
    print("length of dflist(number of group): {}".format(len(dflist)))
    return dflist


def adjustalldf(dflist):
    # function: remove nan rows, move sunken column up, remove header except for group 1
    # input: dflist containing dirty df_group
    # output: dflist containing slightly clean df_group
    print("Adjusting columns...")
    for i, dk in enumerate(dflist):
        dk['0'] = dk['1']  # reassign posid column to the first column
        dk['0'] = separate_posid(dk['0'])
        dk['1'] = risecolumn(dk['1'])
        dk['3'] = risecolumn(dk['3'])
        dk['4'] = risecolumn(dk['4'])
        dk = dk.dropna(how='all')
        if i == 0:
            pass  # dont remove header if group first
        else:
            dk = dk.drop([0], axis=0)
        dk = reindex_row(dk)
        dflist[i] = dk
    return dflist


def addcolumn_pos_sum(df):
    column_60 = []
    posSum_cols = ['2', '9', '11', '13', '15', '17', '19']
    for i, row in df.iterrows():
        text_sum = ''
        for j, element in enumerate(posSum_cols):
            el = df.iloc[i][element]
            if(el == ''):
                pass
            else:
                el += ', '
                text_sum += el
        if(text_sum[-2:] == ', '):
            text_sum = text_sum[:-2]
        column_60.append(text_sum)
    df.insert(60, '60', column_60, True)
    return df


def addcolumn_superior(df):
    df_staff = df[['0', '27', '28']].copy()
    df_staff.rename(columns={'0': '7', '27': '61', '28': '62'}, inplace=True)
    newdf = pd.merge(df, df_staff, on='7', how='left').fillna('')
    return newdf


def create_columnName_list(df):
    list_colname = []
    for i, columnName in df.iteritems():
        list_colname.append(columnName[0])
    list_colname.extend(['Pos SUM', 'Superior ID', 'Superior Name'])
    return list_colname


def arrange_renameCol(df, list_columnname):
    index_order = [27, 28, 43, 44, 45, 46, 47, 48, 49, 0, 60, 7, 61, 62, 2, 8, 9, 10, 11, 12,
                   13, 14, 15, 16, 17, 18, 19, 20, 41, 42, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 1, 3, 4]
    column_order = []
    column_name = []
    for i in index_order:
        s = str(i)
        column_order.append(s)  # basically same as index order but as string
        column_name.append(list_columnname[i])
    column_date = [58, 59, 1]
    for col in column_date:
        c = str(col)
        for i in df.index:
            date = df.iloc[i][c]
            date = date.replace('.', '/')
            df.at[i, c] = date
    df = df.reindex(columns=column_order)
    df = df.sort_values(by=['0'], ascending=True)  # takde pun takpe
    df.columns = column_name
    return df


def remove_header(df):
    df = df.drop([0])  # drop row of column name
    df = df.reset_index(drop=True)
    df = df.fillna("")
    return df


mainpath = os.getcwd()
input_path = r'D:\Documents\Python\project-2\database\input\zhplac.xlsx'
# input_path = os.path.join(mainpath, "database/input/", "zhplac.xlsx")
df = pd.read_excel(input_path, skiprows=4, nrows=None)
output_path = os.path.join(mainpath, "database/output/", "cleanzhpla.xlsx")

df = reindex_column(df)  # reindexed column by number
dflist = slicebigdf(df)
dflist = adjustalldf(dflist)
bigdf = pd.concat(dflist).reset_index(drop=True)

list_colname = create_columnName_list(bigdf)  # list of column name
bigdf = remove_header(bigdf)  # for easy processing
bigdf = addcolumn_pos_sum(bigdf)
bigdf = addcolumn_superior(bigdf)
bigdf = arrange_renameCol(bigdf, list_colname)

print(bigdf)
print("Writing to excel...")
bigdf.to_excel(output_path, index=None)
# bigdf.to_csv('bigdf2.csv', index=None)
print("DONE")
