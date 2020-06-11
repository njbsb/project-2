import pandas as pd
import os
import time
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
        if not i == 0:  # dont remove header if group first
            dk = dk.drop([0], axis=0)
        dk = reindex_row(dk)
        dflist[i] = dk
    return dflist


def set_header(df, columnName):
    df = df.drop([0]).reset_index(drop=True)  # drop row of column name
    df.columns = columnName
    return df


def addcolumn_possum(df):
    column_possum = []
    posSum_cols = ['Position', 'Section', 'Department',
                   'Division', 'Sector', 'Comp. Position', 'Business']
    for i, row in df.iterrows():
        text_sum = ''
        for element in posSum_cols:
            el = df.loc[i, element]
            # print(el, type(el))
            if(el == '' or type(el) is float):
                pass
            else:
                el += ', '
                text_sum += el
        if(text_sum[-2:] == ', '):
            text_sum = text_sum[:-2]
        column_possum.append(text_sum)
    df['Pos SUM'] = column_possum
    return df


def addcolumn_superior(df):
    df_staff = df[['Pos ID', 'Staff No', 'Staff Name']].copy()
    df_staff.rename(columns={'Pos ID': 'Superior Pos ID',
                             'Staff No': 'Superior ID', 'Staff Name': 'Superior Name'}, inplace=True)
    newdf = pd.merge(df, df_staff, on='Superior Pos ID', how='left').fillna('')
    return newdf


def arrange_column(df):
    col_order = ['Staff No', 'Staff Name', 'SG', 'Lvl', 'Tier', 'JG', 'Est.JG', 'EQV JG', 'Conso JG', 'OLIVE JG', 'Pos ID', 'Pos SUM', 'Superior Pos ID', 'Superior ID', 'Superior Name', 'Position', 'Sec ID', 'Section', 'Dept ID', 'Department', 'Division ID', 'Division', 'Sector ID', 'Sector', 'Comp. ID Position', 'Comp. Position', 'Buss ID', 'Business', 'C.Center', 'Gender', 'Race', 'Zone', 'Home SKG', 'Pos.SKG', 'JobID(C)', 'Level', 'NE Area', 'Loc ID', 'Location Text', 'Vac Date', 'Start Date', 'End Date', 'Vacancy Status', 'Email Address', 'Mirror Pos ID', 'Mirror Position', 'EG Post. ID', 'EG Position', 'ESG Post. ID', 'ESG Position', 'PT1', 'PT2', 'Overseas Staff No', 'Overseas Staff Name', 'Overseas Staff Comp. ID', 'Overseas Staff Comp.', 'Staff Comp. ID', 'Staff Comp.', 'Staff EG ID', 'Staff EG', 'Staff ESG ID', 'Staff ESG', 'Staff Work Contract ID', 'Staff Work Contract'
                 ]
    # colname_date = ['End Date', 'Vac Date', 'Start Date']
    # for colname, column in df.iteritems():
    #     if colname in colname_date:
    #         for i, row in column.iteritems():
    #             df.loc[i, colname] = row.replace('.', '/')
    #     elif colname == 'Email Address':
    #         for i, row in column.iteritems():
    #             df.loc[i, colname] = row.lower()
    df = df.reindex(columns=col_order)
    return df


def addcolumn_consoJG(df):
    column_consoJG = []
    jg_columns = ['JG', 'Est.JG', 'EQV JG']

    for i, row in df.iterrows():
        jg = df.loc[i, 'JG']
        estjg = df.loc[i, 'Est.JG']
        eqvjg = df.loc[i, 'EQV JG']
        if not jg == '':
            column_consoJG.append(str(jg))
        else:
            if not estjg == '':
                column_consoJG.append('Est. ' + str(estjg))
            else:
                if not eqvjg == '-':
                    column_consoJG.append('Eqv. ' + str(eqvjg))
                else:
                    column_consoJG.append('No JG')

    # print(column_consoJG, len(column_consoJG))
    df['Conso JG'] = column_consoJG
    return df


def clean_rawdf(df):
    df = reindex_column(df)
    dflist = slicebigdf(df)
    dflist = adjustalldf(dflist)
    reformat_df = pd.concat(dflist).reset_index(drop=True)
    return reformat_df


def format_zhpla(df):
    list_colname = [columnName[0] for i, columnName in df.iteritems()]
    df = set_header(df, list_colname)  # for easy processing
    df = addcolumn_possum(df)
    df = addcolumn_superior(df)
    df = addcolumn_consoJG(df)
    print('arranging columns.')
    df = arrange_column(df)
    return df


def set_input_output(opt):
    input_path = r'D:\Documents\Python\project-2\database\input\ZHPLA_June2020.xlsx' if opt == '1' else r'D:\Documents\Python\project-2\database\input\zhplac.xlsx'
    output_path = r'D:\Documents\Python\project-2\database\output\zh_full_june.xlsx' if opt == '1' else r'D:\Documents\Python\project-2\database\output\zh_partial.xlsx'
    return input_path, output_path


mainpath = os.getcwd()
opt = input('1. Full ZH\n2. Partial ZH\nAnswer: ')
input_path, output_path = set_input_output(opt)
starttime = time.time()
rawdf = pd.read_excel(input_path, skiprows=4, nrows=None)
clean_df = clean_rawdf(rawdf)
finaldf = format_zhpla(clean_df)

print(finaldf)
print("Writing to excel...")
finaldf.to_excel(output_path, index=None)
endtime = time.time()
totaltime = int(endtime - starttime)
print('time elapsed: ', str(datetime.timedelta(seconds=totaltime)))
# bigdf.to_csv('bigdf2.csv', index=None)
print("DONE")
