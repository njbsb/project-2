import pandas as pd
import time
import os
import datetime


def logDuration(start, end):
    totaltime = int(end - start)
    return str(datetime.timedelta(seconds=totaltime))


def calculate_age(born):
    born = datetime.datetime.strptime(born, "%d.%m.%Y").date()
    today = datetime.date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))


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
    ppa5yrs_col = ['PPA-15', 'PPA-16', 'PPA-17', 'PPA-18', 'PPA-19']
    df['PPA'] = df[ppa5yrs_col].apply(
        lambda x: ', '.join(x.dropna().astype(str)), axis=1)
    return df


def addcolumn_SGym(df):
    df[['SG Since', 'SG (YM)']] = df['SG Since'].apply(
        lambda x: pd.Series(str(x).split(";")))
    df['SG Since'] = df['SG Since'].replace(
        to_replace='nan', value=float('nan'))
    return df


def addcolumn_age(df):
    df['Age'] = df['Date of Birth'].apply(
        lambda x: calculate_age(x) if not pd.isna(x) else x)
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
    for i in range(count):
        start = listofindex[i]
        if i < count-1:
            end = listofindex[i + 1]
            df_ = df.iloc[start + gap: end]
        else:
            df_ = df.iloc[start + gap:]
        for i in range(2):
            df_ = df_.dropna(how='all', axis=i)
        df_.columns = df_.iloc[0]
        df_ = df_[1:]
        dflist.append(df_)
    print('group =', len(dflist))
    endtime = time.time()
    print('getDFGroup:', logDuration(starttime, endtime))
    return dflist


def getCleanDF(dflist):
    count = 0
    for df in dflist:
        count += len(df)
    cleandf = pd.concat(dflist).reset_index(drop=True)
    print('count =', count)
    return cleandf


def arrange_column(df):
    start = time.time()
    col_order = ['Personnel Number', 'Formatted Name of Employee or Applicant', 'Gender Key Desc', 'Age', 'Sal. Grade', 'Lvl',
                 'Job Grade', 'PPA-19', 'PPA-18', 'PPA-17', 'PPA-16', 'PPA-15', 'PPA', 'Date of Birth', 'Country of Birth',
                 'Country of Birth Desc', 'State', 'State Desc', 'Birthplace', 'Nationality', 'Nationality Desc', 'Position',
                 'Section', 'Department', 'Division', 'Sector', 'Company', 'SG Since', 'SG (YM)', 'Join Date', 'Date Post',
                 'Join Dept', 'Join Div.', 'Join Comp.', 'Pos. SKG', 'Home SKG', 'Work Center', 'Task Type', 'Task Type Desc', 'Date of Task']
    df = df.drop(['Cluster'], axis=1)
    df = df.reindex(columns=col_order)

    colname_date = ['Date of Birth', 'Join Dept',
                    'SG Since', 'Date Post', 'Join Div.', 'Join Comp.', 'Join Date', 'Date of Task']
    for colname in colname_date:
        df[colname] = df[colname].str.replace('.', '/', regex=False)

    colsemicolon = ['Position', 'Section',
                    'Department', 'Division', 'Work Center']
    for colname in colsemicolon:
        df[colname] = df[colname].str.replace(';', ',', regex=False)

    df.rename(columns={'Personnel Number': 'Staff Number',
                       'Formatted Name of Employee or Applicant': 'Staff Name',
                       'Gender Key Desc': 'Gender',
                       'Sal. Grade': 'SG',
                       'Job Grade': 'JG',
                       'Date Post': 'Date in Position',
                       'Join Dept': 'Date in Department',
                       'Join Div.': 'Date in Division',
                       'Join Comp.': 'Date in Company',
                       'Join Date': 'Date in PETRONAS'}, inplace=True)
    # df.dropna(subset=['Staff Number'])
    df = df[df['Staff Number'].notna()]
    end = time.time()
    print('arranged columns, ', logDuration(start, end))
    return df


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
    df = arrange_column(df)
    print('FINAL\n', df)
    return df


def postProcess(df, outputfile):
    start = time.time()
    df.to_excel(outputfile, index=None)
    os.startfile(outputfile)
    end = time.time()
    print('writing to file, ', logDuration(start, end))


# zpdevpath = r'D:\Documents\Python\project-2\database\input\zpdevc.xlsx'
zpdevpath = r'D:\Documents\Python\project-2\database\input\raw_zpdev_june2020.xlsx'

try:
    outputname = preProcess()
    start = time.time()
    df = mainProcess(zpdevpath)
    postProcess(df, outputname)
except Exception as e:
    print('Process failed:', e)
finally:
    end = time.time()
    print('time elapsed:', logDuration(start, end))
