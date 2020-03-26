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
        df_ = df_.drop(['4'], axis=1)  # drop the empty column at index 4
        df_ = reindex_row(df_)
        df_ = reindex_column(df_)
        dflist.append(df_)
    print("length of dflist(number of group): {}".format(len(dflist)))
    return dflist


def adjustalldf(dflist):
    print("Adjusting columns...")
    for i, dk in enumerate(dflist):
        dk['0'] = dk['1']  # reassign posid column to the first column
        d0 = dk['0']
        d1 = dk['1']
        d3 = dk['3']
        d4 = dk['4']
        dk['0'] = separate_posid(d0)
        dk['1'] = risecolumn(d1)
        dk['3'] = risecolumn(d3)
        dk['4'] = risecolumn(d4)
        dk = dk.dropna(how='all')
        if i == 0:
            pass
        else:
            dk = dk.drop([0], axis=0)
        # rearrange rows
        dk = reindex_row(dk)
        dflist[i] = dk
    return dflist


# CHANGE INPUT PATH 'zhpla_path' to 'file_name.xlsx'
mainpath = os.getcwd()
zhpla_path = os.path.join(mainpath, "database/input/", "zhpla-march2020.xlsx")
df = pd.read_excel(zhpla_path, skiprows=4, nrows=None)

# returns reindexed big df
df = reindex_column(df)

# returns list that contains index to use for slicing, and its length
listofindex, lenlist = getindexlist(df)

# returns list of df that has been sliced into groups
dflist = slicebigdf(df, lenlist, listofindex)

# return list of df that has been adjusted its columns
dflist = adjustalldf(dflist)

cleanzhpla_path = os.path.join(mainpath, "database/output/", "cleanzhpla.xlsx")

# combines all df in the dflist
print("Concatenating all dataframe...")
bigdf = pd.concat(dflist)

# CHANGE OUTPUT PATH 'cleanzhpla_path' to 'file_name.xlsx'
print("Writing to excel...")
bigdf.to_excel(cleanzhpla_path, index=None, header=None)
print("DONE")
