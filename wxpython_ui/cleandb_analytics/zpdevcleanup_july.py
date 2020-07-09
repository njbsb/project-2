import pandas as pd
import time
import os
import sys
import datetime
import numpy as np
from dateutil.relativedelta import relativedelta


def logDuration(start, end):
    totaltime = int(end - start)
    return str(datetime.timedelta(seconds=totaltime))


def calculate_age(born):
    born = datetime.datetime.strptime(born, "%d.%m.%Y").date()
    today = datetime.date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))


def date_integrity(datetime_row):
    try:
        if pd.to_datetime(datetime_row, format="%d.%m.%Y"):
            return datetime_row
            # return pd.to_datetime(datetime_row, format="%d.%m.%Y")
    except ValueError:
        return np.nan


def calculate_durationYM(date):
    if not pd.isnull(date):
        # print(date, date.date())
        dip = date.date()
        rdelta = relativedelta(datetime.date.today(), dip)
        dateYM = str(rdelta.years) + "y" + str(rdelta.months) + "m"
        return dateYM
    else:
        return pd.NaT
    # return dateYM


def reindex_row(df):
    rowlength = df.shape[0]
    newindex = pd.Series([i for i in range(rowlength)])
    df = df.set_index(newindex)
    return df


def reindex_column(df):
    colname = list(df.columns)
    colnames = [i.replace(i, str(colname.index(i))) for i in colname]
    df.columns = colnames
    return df


def addcolumn_ppa5yrs(df):
    ppa5yrs_col = ['PPA-19', 'PPA-18', 'PPA-17', 'PPA-16', 'PPA-15']
    df_ppa = df[ppa5yrs_col]

    df['Cluster'] = df['Cluster'].replace(
        {'High': 'H', 'Solid': 'S', 'Low': 'L', '-': ''})

    for eachname, eachcolumn in df_ppa.iteritems():
        currentindex = df.columns.get_loc(eachname)
        clusterindex = currentindex + 1
        df_eachppa = df.iloc[:, currentindex: clusterindex + 1]
        # df[eachname] = df_eachppa.apply(
        #     lambda x: ''.join(x.dropna().astype(str)), axis=1)
        df[eachname] = df_eachppa[eachname].astype(
            str) + df_eachppa['Cluster'].astype(str)
        # print(df[eachname])
    # print(df[ppa5yrs_col])
    df['PPA'] = df[ppa5yrs_col].apply(
        lambda x: ', '.join(x.dropna().astype(str)), axis=1)
    df = df.drop(['Cluster'], axis=1)
    # print(df['PPA'])
    return df


def addcolumn_SGym(df):
    df[['SG Since', 'SG (YM)']] = df['SG Since'].apply(
        lambda x: pd.Series(str(x).split(";")))
    df['SG Since'] = df['SG Since'].replace(
        to_replace='nan', value=float('nan'))
    return df


def addcolumn_age(df):
    # try:
    df['Date of Birth'] = df['Date of Birth'].apply(date_integrity)
    df['Age'] = df['Date of Birth'].apply(
        lambda x: calculate_age(x) if not pd.isna(x) else x)
    # except Exception as e:
    #     print('Process failed:', e)
    #     exc_type, exc_obj, exc_tb = sys.exc_info()
    #     fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    #     print(exc_type, fname, exc_tb.tb_lineno)
    return df


def addcolumn_yearIn(df):
    starttime = time.time()
    dateIn = ['Date Post', 'Join Date']
    yearIn = ['YiPosition', 'YiPETRONAS']
    df_dateIn = df[dateIn]
    for date, yearmonth in zip(dateIn, yearIn):
        # df_dateIn[date] = df_dateIn[date].apply(date_integrity)
        df_dateIn[date] = pd.to_datetime(
            df_dateIn[date], format="%d.%m.%Y", errors='coerce')
        df_dateIn[yearmonth] = df_dateIn[date].apply(
            lambda x: calculate_durationYM(x))
        df[yearmonth] = df_dateIn[yearmonth]
    # print(df_dateIn)
    endtime = time.time()
    print('yearIn,', logDuration(starttime, endtime))
    return df


def arrange_column(df):
    start = time.time()
    colname_date = ['Date of Birth', 'Join Dept',
                    'SG Since', 'Date Post', 'Join Div.', 'Join Comp.', 'Join Date', 'Date of Task']
    for colname in colname_date:
        df[colname] = df[colname].str.replace('.', '/', regex=False)
    colsemicolon = ['Position', 'Section', 'Sector',
                    'Department', 'Division', 'Work Center']
    for colname in colsemicolon:
        df[colname] = df[colname].str.replace(';', ',', regex=False)

    col_order = ['Personnel Number', 'Formatted Name of Employee or Applicant', 'Gender Key Desc', 'Age', 'Sal. Grade', 'Lvl',
                 'Job Grade', 'PPA-19', 'PPA-18', 'PPA-17', 'PPA-16', 'PPA-15', 'PPA', 'Date of Birth', 'Country of Birth',
                 'Country of Birth Desc', 'State', 'State Desc', 'Birthplace', 'Nationality', 'Nationality Desc', 'Position',
                 'Section', 'Department', 'Division', 'Sector', 'Company', 'SG Since', 'SG (YM)', 'Join Date', 'YiPETRONAS', 'Date Post', 'YiPosition',
                 'Join Dept', 'Join Div.', 'Join Comp.', 'Pos. SKG', 'Home SKG', 'Work Center', 'Task Type', 'Task Type Desc', 'Date of Task']
    df = df.reindex(columns=col_order)
    df = df[df['Personnel Number'].notna()]
    # df.rename(columns={'Personnel Number': 'Staff Number',
    #                    'Formatted Name of Employee or Applicant': 'Staff Name',
    #                    'Gender Key Desc': 'Gender',
    #                    'Sal. Grade': 'SG',
    #                    'Job Grade': 'JG',
    #                    'Date Post': 'Date in Position',
    #                    'Join Dept': 'Date in Department',
    #                    'Join Div.': 'Date in Division',
    #                    'Join Comp.': 'Date in Company',
    #                    'Join Date': 'Date in PETRONAS'}, inplace=True)
    # df.dropna(subset=['Staff Number'])

    end = time.time()
    print('arranged columns, ', logDuration(start, end))
    return df


def getGroup(df):
    # list contain index of new group of data, detects 'zhpla report' from column 0/A
    listofindex = []
    for i, row in df.iterrows():
        if type(row[0]) == str and 'HR' in row[0]:
            listofindex.append(i)
    return listofindex, len(listofindex)


def getDFGroup(df):
    starttime = time.time()
    listofindex, count = getGroup(df)
    dflist = []
    gap = 5
    dropaxis = [0, 1]
    for i in range(count):
        start = listofindex[i]
        if i < count-1:
            end = listofindex[i + 1]
            df_ = df.iloc[start + gap: end]
        else:
            df_ = df.iloc[start + gap:]
        for axis in dropaxis:
            df_ = df_.dropna(how='all', axis=axis)
        dflist.append(df_)
    print('group =', len(dflist))
    endtime = time.time()
    print('getDFGroup:', logDuration(starttime, endtime))
    return dflist


def getCleanDF(dflist):
    count = 0
    for i, df in enumerate(dflist):
        df = reindex_row(df)

        df.columns = df.iloc[0]
        df = df[1:]
        dflist[i] = df
        count += len(df)
    cleandf = pd.concat(dflist).reset_index(drop=True)
    # print('cleandf\n', cleandf)
    # print('count =', count)
    return cleandf


def preProcess():
    output = input(
        'Please insert output file name (without .xlsx)\nOutput file name: ')
    outputfile = output + '.xlsx'
    return outputfile


def mainProcess(zpdevpath):
    dfRaw = pd.read_excel(zpdevpath, skiprows=4)
    dflist = getDFGroup(dfRaw)
    cleandf = getCleanDF(dflist)

    df = addcolumn_ppa5yrs(cleandf)
    df = addcolumn_SGym(df)
    df = addcolumn_age(df)
    df = addcolumn_yearIn(df)
    df = arrange_column(df)
    print('FINAL\n', df)
    return df


def postProcess(df, outputfile):
    start = time.time()
    # df.to_excel(outputfile, index=None)
    # os.startfile(outputfile)
    end = time.time()
    print('writing to file, ', logDuration(start, end))


# zpdevpath = r'D:\Documents\Python\project-2\database\input\zpdevc.xlsx'
zpdevpath = r'D:\Documents\Python\project-2\database\input\raw_zpdev_june2020.xlsx'

# try:
#     outputname = preProcess()
#     start = time.time()
#     df = mainProcess(zpdevpath)
#     postProcess(df, outputname)
# except Exception as e:
#     print('Process failed:', e)
# finally:
#     end = time.time()
#     print('time elapsed:', logDuration(start, end))
