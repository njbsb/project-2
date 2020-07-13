import pandas as pd
import numpy as np
import os
import sys
import time
import datetime
# import zpdevcleanup_july as zp
from dateutil.relativedelta import relativedelta


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


def getGroup(df):
    # list contain index of new group of data, detects 'zhpla report' from column 0/A
    try:
        listofindex = []
        for i, row in df.iterrows():
            if type(row[0]) == str and 'HR' in row[0]:
                listofindex.append(i)
        return listofindex, len(listofindex)
    except Exception as e:
        print(e)


def getDFGroup(df):
    try:
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
        return dflist
    except Exception as e:
        print(e)
    finally:
        print('getDFGroup:', logDuration(starttime, time.time()))
    # return dflist


def getCleanDF(dflist):
    try:
        count = 0
        for i, df in enumerate(dflist):
            df = reindex_row(df)

            df.columns = df.iloc[0]
            df = df[1:]
            dflist[i] = df
            count += len(df)
        cleandf = pd.concat(dflist).reset_index(drop=True)
        return cleandf
    except Exception as e:
        print(e)
        return None


def addcolumn_ppa5yrs(df):
    try:
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
        df['Date of Birth'] = df['Date of Birth'].apply(date_integrity)
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
    except Exception as e:
        print(e)
    finally:
        return df


def arrange_column(df):
    try:
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
        # df.dropna(subset=['Staff Number'])
    except Exception as e:
        print(e)
    finally:
        # end = time.time()
        print('arranged columns, ', logDuration(start, time.time()))
        return df


def mainProcess(zpdevmain_path, zpdevretire_path, outputname):
    df_main = pd.read_excel(zpdevmain_path)
    df_retire = pd.read_excel(zpdevretire_path)

    dflist = [df_main, df_retire]
    for i, df in enumerate(dflist):
        df_group = getDFGroup(df)
        cleandf = getCleanDF(df_group)
        dflist[i] = cleandf

    df_main, df_retire = dflist
    combined_zpdev = pd.merge(df_main, df_retire[[
        'Personnel Number', 'Task Type', 'Task Type Desc', 'Date of Task']], on='Personnel Number', how='left')
    combined_zpdev = addcolumn_ppa5yrs(combined_zpdev)
    combined_zpdev = addcolumn_SGym(combined_zpdev)
    combined_zpdev = addcolumn_age(combined_zpdev)
    combined_zpdev = addcolumn_yearIn(combined_zpdev)
    combined_zpdev = arrange_column(combined_zpdev)

    combined_zpdev.to_excel(outputname, index=None)
    return combined_zpdev


zpdev_main_path = r'D:\Documents\Python\project-2\database\input\July\ZPDEV_July_2020.xlsx'
zpdev_retire_path = r'D:\Documents\Python\project-2\database\input\July\Retirement_july2020.xlsx'
outputname = 'combined_zpdevjuly.xlsx'

combined_zpdev = mainProcess(zpdev_main_path, zpdev_retire_path, outputname)
print(combined_zpdev)


# ______________________________________________________
# UNCOMMENT LINE BELOW TO  OPEN FILE AFTER EXECUTION
# ______________________________________________________
# os.startfile(outputname)
