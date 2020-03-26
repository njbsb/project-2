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


def slicebigdf(df, count, listofindex):
    dflist = []
    for i in range(count):
        st = listofindex[i]  # st is the starting index of the group
        if i < lenlist-1:  # if st is not the last element in the list, then we can set the endpoint
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


mainpath = os.getcwd()
zpdev_path = os.path.join(mainpath, "database/input/", "zpdev-march2020.xlsx")

# CHANGE INPUT PATH 'zhpla_path' to 'file_name.xlsx'
zpdev_file = pd.ExcelFile('c:/syafiq.rusla.....xlsx')

df = pd.read_excel(zpdev_path, skiprows=4, nrows=None)

# returns reindexed big df
df = reindex_column(df)

# returns list that contains index to use for slicing, and its length
listofindex, lenlist = getindexlist(df)

# returns list of df that has been sliced into groups
dflist = slicebigdf(df, lenlist, listofindex)

# return list of df that has been adjusted its columns
dflist = adjustalldf(dflist)

cleanzpdev_path = os.path.join(mainpath, "database/output/", "cleanzpdev.xlsx")

# combines all df in the dflist
print("Concatenating all dataframe...")
bigdf = pd.concat(dflist)

# CHANGE OUTPUT PATH 'cleanzhpla_path' to 'file_name.xlsx'
print("Writing to excel...")
bigdf.to_excel(cleanzpdev_path, index=None, header=None)
print("DONE")
