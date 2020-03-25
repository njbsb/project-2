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


mainpath = os.getcwd()
zhpla_path = os.path.join(mainpath, "database/input/", "zhplac.xlsx")
# change zhpla_path to file path
zhpla_file = pd.ExcelFile(zhpla_path)

df = pd.read_excel(zhpla_path, skiprows=4, nrows=None)
df = reindex_column(df)

listofi = []  # list contain index of new group of data, detects 'zhpla report' from column 0/A
for i, row in df.iterrows():
    if(pd.isna(row[0])):
        pass
    else:
        listofi.append(i)
print("length of listofi: {}".format(len(listofi)))

dflist = []
lenlist = len(listofi)
for i in range(lenlist):
    st = listofi[i]  # st is the starting index of the group
    if i < lenlist-1:  # if st is not the last element in the list, then we can set the endpoint
        en = listofi[i+1]  # endpoint is the next element(id)
        df_ = df.iloc[st+3:en]  # slice the big df into the specified range
    else:  # if st is the last element
        df_ = df.iloc[st+3:]  # no need to set the endpoint
    df_ = df_.drop(['4'], axis=1)  # drop the empty column at index 4
    df_ = reindex_row(df_)
    df_ = reindex_column(df_)
    dflist.append(df_)
print("length of dflist: {}".format(len(dflist)))

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
    dk = reindex_row(dk)
    dflist[i] = dk

bigdf = pd.concat(dflist)
bigdf.to_excel("zhplap_o1.xlsx", index=None, header=None)
