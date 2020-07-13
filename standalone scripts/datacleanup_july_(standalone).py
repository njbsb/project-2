import pandas as pd
import numpy as np
import time
import os
import sys
import datetime
from dateutil.relativedelta import relativedelta

mydate = datetime.datetime.now()
current_month = mydate.strftime("%B_%Y")
print(current_month)


def reindex_column(df):
    colname = list(df.columns)
    colnames = [i.replace(i, str(colname.index(i))) for i in colname]
    df.columns = colnames
    return df


def reindex_row(df):
    rowlength = df.shape[0]
    newindex = pd.Series([i for i in range(rowlength)])
    df = df.set_index(newindex)
    return df


def calculate_durationYM(date):
    if not pd.isnull(date):
        dip = date.date()
        rdelta = relativedelta(datetime.date.today(), dip)
        dateYM = str(rdelta.years) + "y" + str(rdelta.months) + "m"
        return dateYM
    else:
        return pd.NaT


def calculate_age(born):
    # born = datetime.datetime.strptime(born, "%d.%m.%Y").date()
    born = born.date()
    today = datetime.date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))


def date_integrity(datetime_row):
    try:
        if pd.to_datetime(datetime_row, format="%d.%m.%Y"):
            return datetime_row
            # return pd.to_datetime(datetime_row, format="%d.%m.%Y")
    except ValueError:
        return np.nan


def logDuration(start, end):
    totaltime = int(end - start)
    return str(datetime.timedelta(seconds=totaltime))


def getGroup(df, keyword):
    # list contain index of new group of data, detects keyword from column 0/A
    listofindex = []
    for i, row in df.iterrows():
        if type(row[0]) == str and keyword in row[0]:
            listofindex.append(i)
    return listofindex, len(listofindex)


def getDFGroup(df, keyword):
    gap = 3 if keyword == 'ZHPLA' else 5
    dftype = 'ZHPLA' if keyword == 'ZHPLA' else 'ZPDEV'
    # starttime = time.time()
    listofindex, count = getGroup(df, keyword)
    dflist = []

    for i in range(count):
        start = listofindex[i]
        if i < count-1:
            end = listofindex[i + 1]
            df_ = df.iloc[start + gap: end]
        else:
            df_ = df.iloc[start + gap:]
        for axis in [0, 1]:
            df_ = df_.dropna(how='all', axis=axis)
        dflist.append(df_)

    print('\nNew Dataframe: {}\ngroup:'.format(dftype), len(dflist))
    # print('getDFGroup,', logDuration(starttime, time.time()))
    return dflist


def getCleanDF(dflist, keyword):
    count = 0
    for i, df in enumerate(dflist):
        df = reindex_row(df)

        if keyword == 'ZHPLA':
            df = reindex_column(df)
            for colname, column in df.iteritems():
                if not pd.isna(column[1]):  # if row 1 has value/column name
                    if pd.isna(column[0]):  # if row 0 already has column name
                        df[colname] = df[colname].shift(-1)
                    else:
                        newcol = str(len(df.columns) + 1)
                        df[newcol] = column
                        df[newcol] = df[newcol].shift(-1)  # shift the sunken
                        for dfcol in [df[colname], df[newcol]]:
                            # print("dfcol ", dfcol)
                            for index, row in dfcol.iteritems():
                                if not index == 0:
                                    if index in (1, 2) or index % 2 == 0:
                                        dfcol[index] = float('nan')
            df = df.dropna(how='all')

        df.columns = df.iloc[0]
        df = df[1:]
        dflist[i] = df
        count += len(df)
    cleandf = pd.concat(dflist).reset_index(drop=True)
    print('total row count:', count)
    return cleandf


def correctFormat(df, keyword):
    if keyword == 'ZHPLA':
        colname_date = ['End Date', 'Vac Date', 'Start Date']
        df['Email Address'] = df['Email Address'].str.lower()
    else:
        colname_date = ['Date of Birth', 'Join Dept', 'SG Since', 'Date Post',
                        'Join Div.', 'Join Comp.', 'Join Date', 'Date of Task']
    df = df.replace(';', ',', regex=True)
    df[colname_date] = df[colname_date].apply(
        pd.to_datetime, format='%d.%m.%Y', errors='coerce')
    return df


def addcolumn_possum(df):
    start = time.time()
    possum_column = ['Position', 'Section', 'Department',
                     'Division', 'Sector', 'Comp. Position', 'Business']
    df['Pos SUM'] = df[possum_column].apply(
        lambda x: ', '.join(x.dropna().astype(str)), axis=1)
    print('added column POSSUM, ', logDuration(start, time.time()))
    return df


def addcolumn_superior(df):
    start = time.time()
    df_staff = df[['Pos ID', 'Staff No', 'Staff Name']].drop_duplicates()
    df_staff.rename(columns={'Pos ID': 'Superior Pos ID',
                             'Staff No': 'Superior ID', 'Staff Name': 'Superior Name'}, inplace=True)
    newdf = pd.merge(df, df_staff, on='Superior Pos ID', how='left')
    end = time.time()
    print('added column superior: ', logDuration(start, end))
    return newdf


def addcolumn_consoJG(df):
    start = time.time()
    column_consoJG = []
    jg_head = ['', 'Est. ', 'Eqv. ']
    df_jg = df[['JG', 'Est.JG', 'EQV JG']].fillna('-')
    for i, eachrow in df_jg.iterrows():
        for index, val in enumerate(list(eachrow)):
            if not val == '-':
                conso = jg_head[index] + str(val)
                column_consoJG.append(conso)
                break
            else:
                if(index == 2 and val == '-'):
                    column_consoJG.append('No JG')
                else:
                    continue
    df['Conso JG'] = column_consoJG
    end = time.time()
    print('added column consoJG, ', logDuration(start, end))
    return df


def addcolumn_ppa5yrs(df):
    try:
        ppa5yrs_col = ['PPA-19', 'PPA-18', 'PPA-17', 'PPA-16', 'PPA-15']
        df_ppa = df[ppa5yrs_col].fillna('')
        df['Cluster'] = df['Cluster'].replace(
            {'High': 'H', 'Solid': 'S', 'Low': 'L', '-': '', 'N/A': ''})
        df['Cluster'] = df['Cluster']
        for eachname, eachcolumn in df_ppa.iteritems(): # replace with df_ppa.columns?
            currentindex = df.columns.get_loc(eachname)
            clusterindex = currentindex + 1
            df_eachppa = df.iloc[:, currentindex: clusterindex + 1].fillna('')
            # df[eachname] = df_eachppa.apply(
            #     lambda x: ''.join(x.dropna().astype(str)), axis=1)
            df[eachname] = df_eachppa[eachname].astype(
                str) + df_eachppa['Cluster'].astype(str)

        df['PPA'] = df[ppa5yrs_col].apply(
            lambda x: ', '.join(x.dropna().astype(str)), axis=1)
        df = df.drop(['Cluster'], axis=1)
    except Exception as e:
        print(e)
    finally:
        return df


def addcolumn_SGym(df):
    try:
        df[['SG Since', 'SG (YM)']] = df['SG Since'].apply(
            lambda x: pd.Series(str(x).split(";")))
        df['SG Since'] = df['SG Since'].replace(
            to_replace='nan', value=float('nan'))
    except Exception as e:
        print(e)
    finally:
        return df


def addcolumn_age(df):
    try:
        df['Date of Birth'] = df['Date of Birth'].apply(
            date_integrity)  # return nan if not match date type

        df['Age'] = df['Date of Birth'].apply(
            lambda x: calculate_age(x) if not pd.isna(x) else x)

    except Exception as e:
        print('Process failed:', e)
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
    finally:
        return df


def addcolumn_yearIn(df):
    try:
        dateIn = ['Date Post', 'Join Date']
        yearIn = ['YiPosition', 'YiPETRONAS']
        for date, yearmonth in zip(dateIn, yearIn):
            df[yearmonth] = df[date].apply(
                lambda x: calculate_durationYM(x))
    except Exception as e:
        print(e)
    finally:
        return df


def arrange_column(df, keyword):
    if keyword == 'ZHPLA':
        col_order = ['Staff No', 'Staff Name', 'SG', 'Lvl', 'Tier', 'JG', 'Est.JG', 'EQV JG', 'Conso JG', 'OLIVE JG',
                     'Pos ID', 'Pos SUM', 'Superior Pos ID', 'Superior ID', 'Superior Name', 'Position', 'Sec ID',
                     'Section', 'Dept ID', 'Department', 'Division ID', 'Division', 'Sector ID', 'Sector', 'Comp. ID Position',
                     'Comp. Position', 'Buss ID', 'Business', 'C.Center', 'Gender', 'Race', 'Zone', 'Home SKG', 'Pos.SKG', 'JobID(C)',
                     'Level', 'NE Area', 'Loc ID', 'Location Text', 'Vac Date', 'Start Date', 'End Date', 'Vacancy Status',
                     'Email Address', 'Mirror Pos ID', 'Mirror Position', 'EG Post. ID', 'EG Position', 'ESG Post. ID', 'ESG Position',
                     'PT1', 'PT2', 'Overseas Staff No', 'Overseas Staff Name', 'Overseas Staff Comp. ID', 'Overseas Staff Comp.',
                     'Staff Comp. ID', 'Staff Comp.', 'Staff EG ID', 'Staff EG', 'Staff ESG ID', 'Staff ESG', 'Staff Work Contract ID', 'Staff Work Contract'
                     ]
    else:  # ZPDEV
        col_order = ['Personnel Number', 'Formatted Name of Employee or Applicant', 'Gender Key Desc', 'Age', 'Sal. Grade', 'Lvl',
                     'Job Grade', 'PPA-19', 'PPA-18', 'PPA-17', 'PPA-16', 'PPA-15', 'PPA', 'Date of Birth', 'Country of Birth',
                     'Country of Birth Desc', 'State', 'State Desc', 'Birthplace', 'Nationality', 'Nationality Desc', 'Position', 'Section',
                     'Department', 'Division', 'Sector', 'Company', 'SG Since', 'SG (YM)', 'Join Date', 'YiPETRONAS', 'Date Post', 'YiPosition',
                     'Join Dept', 'Join Div.', 'Join Comp.', 'Pos. SKG', 'Home SKG', 'Work Center', 'Task Type', 'Task Type Desc', 'Date of Task']
        df = df[df['Personnel Number'].notna()]
    df = df.reindex(columns=col_order)
    return df


def prepareDataframe(df_zhpla, df_zpdev):
    # starttime = time.time()
    df_zpdev.rename(columns={'Pos. SKG': 'Pos. SKG Zpdev',
                             'Home SKG': 'Home SKG Zpdev',
                             'Personnel Number': 'Staff No',
                             'Formatted Name of Employee or Applicant': 'Staff Name',
                             'Gender Key Desc': 'Gender',
                             'Sal. Grade': 'SG',
                             'Join Date': 'Date in PETRONAS',
                             'Date Post': 'Date in Position',
                             'Join Dept': 'Date in Department',
                             'Join Div.': 'Date in Division',
                             'Join Comp.': 'Date in Company',
                             'Date of Task': 'Retirement Date'
                             }, inplace=True)
    # print('prepareDataframe:', logDuration(starttime, time.time()))
    return df_zhpla, df_zpdev


def merge_database(df_zhpla, df_zpdev):
    try:
        starttime = time.time()
        mergeOn = ['Staff No', 'Staff Name']
        df_merge = pd.merge(df_zhpla, df_zpdev, on=mergeOn, how='outer')

        compare_list = [['SG_x', 'SG_y'], ['Lvl_x', 'Lvl_y'], ['Position_x', 'Position_y'], ['Section_x', 'Section_y'], [
            'Department_x', 'Department_y'], ['Division_x', 'Division_y'], ['Sector_x', 'Sector_y'], ['Gender_x', 'Gender_y']]
        for column in compare_list:
            df_merge[column[0]] = df_merge.apply(
                lambda row: row[column[1]] if pd.isnull(row[column[0]]) else row[column[0]], axis=1)
            xcolumn = column[0]
            # print(df_merge[column])
            if(xcolumn[-2:] == '_x'):
                df_merge.rename(columns={xcolumn: xcolumn[:-2]}, inplace=True)

        JG_cl = ['Conso JG', 'Job Grade']
        df_merge[JG_cl[0]] = df_merge.apply(lambda row: row[JG_cl[1]] if pd.isnull(
            row[JG_cl[0]]) or row[JG_cl[0]] == 'No JG' else row[JG_cl[0]], axis=1)

        arrangedcol = ['Staff No', 'Staff Name', 'Gender', 'Race', 'Age', 'SG', 'Lvl', 'Tier', 'Conso JG', 'Pos ID', 'Pos SUM', 'Superior Pos ID', 'Superior ID',
                       'Superior Name', 'SG Since', 'SG (YM)', 'YiPETRONAS', 'YiPosition', 'PPA-19', 'PPA-18', 'PPA-17', 'PPA-16', 'PPA-15', 'PPA', 'Date in PETRONAS',
                       'Date in Position', 'Date in Department', 'Date in Division', 'Date in Company', 'Date of Birth', 'Country of Birth', 'Country of Birth Desc',
                       'State', 'State Desc', 'Birthplace', 'Nationality', 'Nationality Desc', 'Position', 'Sec ID', 'Section', 'Dept ID', 'Department', 'Division ID',
                       'Division', 'Sector ID', 'Sector', 'Comp. ID Position', 'Comp. Position', 'Buss ID', 'Business', 'C.Center', 'Home SKG', 'Pos.SKG', 'Pos. SKG Zpdev',
                       'Home SKG Zpdev', 'JobID(C)', 'Level', 'Loc ID', 'Location Text', 'Work Center', 'Email Address', 'Task Type', 'Task Type Desc', 'Retirement Date']

        merged_columns = list(df_merge.columns)
        columnCheck = all(item in merged_columns for item in arrangedcol)
        if columnCheck:
            print('Talent Analytics: Generated')
            final_analytics = df_merge[arrangedcol]
            return final_analytics
        else:
            print('Talent Analytics: Failed to generate')
            return df_merge

    except Exception as e:
        print(e)
        return None
    finally:
        print('merge_database:', logDuration(starttime, time.time()))


def process_zhpla(zhplapath):
    keyword = 'ZHPLA'
    dfRaw = pd.read_excel(zhplapath)
    dflist = getDFGroup(dfRaw, keyword)
    df = getCleanDF(dflist, keyword)

    df = correctFormat(df, keyword)

    df = addcolumn_possum(df)
    df = addcolumn_superior(df)
    df = addcolumn_consoJG(df)
    df = arrange_column(df, keyword)
    return df


def process_zpdev(zpdevpath, zpdev_retirepath):
    keyword = 'HR'
    df_main = pd.read_excel(zpdevpath)
    df_retire = pd.read_excel(zpdev_retirepath)

    dflist = [df_main, df_retire]

    for i, df in enumerate(dflist):
        df_group = getDFGroup(df, keyword)
        cleandf = getCleanDF(df_group, keyword)
        dflist[i] = cleandf

    df_main, df_retire = dflist
    df_merge = pd.merge(df_main, df_retire[[
        'Personnel Number', 'Task Type', 'Task Type Desc', 'Date of Task']],
        on='Personnel Number', how='left')

    # split this first before apply datetime to column date
    df_merge = addcolumn_SGym(df_merge)
    df_merge = correctFormat(df_merge, keyword)

    df_merge = addcolumn_ppa5yrs(df_merge)
    df_merge = addcolumn_age(df_merge)
    df_merge = addcolumn_yearIn(df_merge)
    df_merge = arrange_column(df_merge, keyword)

    return df_merge


def cleanOnly(data_dict):
    dict_length = len(data_dict)
    if dict_length == 1:
        zhplapath = data_dict['zhpla']
        # do zhpla

    if dict_length == 2:
        zpdevpath = data_dict['zpdevmain']
        zpdev_retirepath = data_dict['zpdevretire']
        # do zpdev

    if dict_length == 3:
        zhplapath = data_dict['zhpla']
        zpdevpath = data_dict['zpdevmain']
        zpdev_retirepath = data_dict['zpdevretire']
        # do zhpla and zpdev

    new_dict = {}

    if 'zhplapath' in locals():
        df_zhpla = process_zhpla(zhplapath)
        new_dict['zhpla'] = df_zhpla

    if 'zpdevpath' in locals() and 'zpdev_retirepath' in locals():
        df_zpdev = process_zpdev(zpdevpath, zpdev_retirepath)
        new_dict['zpdev'] = df_zpdev

    return new_dict


def combineOnly(data_dict):
    if 'zhpla' in data_dict:
        if isinstance(data_dict['zhpla'], pd.DataFrame):
            df_zhpla = data_dict['zhpla']
        else:  # then it is path
            df_zhpla = pd.read_excel(data_dict['zhpla'])
    if 'zpdev' in data_dict:
        if isinstance(data_dict['zpdev'], pd.DataFrame):
            df_zpdev = data_dict['zpdev']
        else:
            df_zpdev = pd.read_excel(data_dict['zpdev'])

    df_zhpla, df_zpdev = prepareDataframe(df_zhpla, df_zpdev)
    df_analytics = merge_database(df_zhpla, df_zpdev)
    return df_analytics


def mainProcess(data_dictionary, operation_type):
    starttime = time.time()
    if int(operation_type) == 1:
        df_dictionary = cleanOnly(data_dictionary)
        # can contain 1 or 2 values either zhpla/zpdev or both
    else:
        df_dictionary = {}
        if int(operation_type) == 2:
            df_analytics = combineOnly(data_dictionary)

        if int(operation_type) == 3:
            df_dictionary = cleanOnly(data_dictionary)
            df_analytics = combineOnly(df_dictionary)

        df_dictionary['analytics'] = df_analytics
    print('Process completed: ', logDuration(starttime, time.time()))
    return df_dictionary


def preProcess(zhplapath, zpdevpath, zpdev_retirepath):
    operation = None
    while operation not in ['1', '2', '3']:
        operation = input(
            '\nA. Select operation:\n1. Clean (1)\n2. Combine (2)\n3. Clean & combine (3)\nAnswer: ')

    if operation == '1':
        file_option = input(
            '\nB. Select files:\n1. ZHPLA (1)\n2. ZPDEV (2)\n3. ZHPLA & ZPDEV (3)\nAnswer: ')
        if file_option == '1':
            data_dict = {'zhpla': zhplapath}
        elif file_option == '2':
            data_dict = {'zpdevmain': zpdevpath,
                         'zpdevretire': zpdev_retirepath}
        elif file_option == '3':
            data_dict = {'zhpla': zhplapath, 'zpdevmain': zpdevpath,
                         'zpdevretire': zpdev_retirepath}

    if operation == '2':
        zhplapath = r'D:\Documents\Python\project-2\database\input\July\cleanzhpla_july.xlsx'
        zpdevpath = r'D:\Documents\Python\project-2\database\input\July\cleanzpdev_july.xlsx'
        data_dict = {'zhpla': zhplapath, 'zpdev': zpdevpath}

    if operation == '3':
        data_dict = {'zhpla': zhplapath, 'zpdevmain': zpdevpath,
                     'zpdevretire': zpdev_retirepath}

    return data_dict, operation


zhplapath = r'D:\Documents\Python\project-2\database\input\July\ZHPLA_July2020.xlsx'
zpdev_mainpath = r'D:\Documents\Python\project-2\database\input\July\ZPDEV_July_2020.xlsx'
zpdev_retirepath = r'D:\Documents\Python\project-2\database\input\July\Retirement_july2020.xlsx'


data_dictionary, operation = preProcess(
    zhplapath, zpdev_mainpath, zpdev_retirepath)

df_dictionary = mainProcess(data_dictionary, operation)
print(len(df_dictionary))
for key, value in df_dictionary.items():
    print(key, '\n', value)
    outputfile = (key + current_month).upper() + '.xlsx'
    value.to_excel(outputfile, index=None)
    os.startfile(outputfile)
