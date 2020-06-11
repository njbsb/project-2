import pandas as pd
import os
from datetime import date, datetime
from dateutil.relativedelta import relativedelta


def date_to_duration(date):
    dip = date.date()
    rdelta = relativedelta(date.today(), dip)
    dateYM = str(rdelta.years) + "y" + str(rdelta.months) + "m"
    return dateYM


def merge_database(df_zh, df_zp):
    print('merging databases')
    mergeOn = ['Staff No', 'Staff Name']
    df_merge = pd.merge(df_zh, df_zp, on=mergeOn, how='outer')

    compare_list = [[11, 53], [13, 58], [15, 57], [17, 56], [
        19, 55], [21, 54], [25, 34], [5, 38], [2, 36], [3, 37]]

    for i, cl in enumerate(compare_list):
        for j, row in df_merge.iloc[:, cl[0]].iteritems():
            if type(row) == float:
                df_merge.iloc[j, cl[0]] = df_merge.iloc[j, cl[1]]
        xcolumn = df_merge.columns[cl[0]]
        if(xcolumn[-2:] == '_x'):
            df_merge.rename(columns={xcolumn: xcolumn[:-2]}, inplace=True)

    drop_column = [df_merge.columns[cl[1]] for cl in compare_list]
    for drop in drop_column:
        df_merge = df_merge.drop(drop, axis=1)
    column = list(df_merge.columns.values)
    ta_column = [0, 1, 25, 26, 34, 2, 3, 4, 5, 6, 7, 8, 9, 10, 49, 50, 62, 63, 35, 36, 37, 38, 39, 40, 51, 52, 53, 54, 55, 41, 42, 43,
                 44, 45, 46, 47, 48, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 27, 28, 56, 57, 29, 30, 31, 32, 58, 33, 59, 60, 61]
    arrangedcol = [column[i] for i in ta_column]
    df_merge = df_merge[arrangedcol]
    df_merge = df_merge.sort_values(by=['Pos ID'], ascending=True)
    add_spcolumn(df_merge)
    print('writing to excel')
    return df_merge


def concat_possum(df):
    # for staff that only exist in zp
    possum_column = [21, 26, 25, 24, 23, 22]
    for i, row in df.loc[:, 'Pos SUM'].iteritems():
        if type(row) == float:
            possum_str = ''
            for cid in possum_column:
                if not type(df.iloc[i, cid]) == float:
                    pass
            # concat the shit
            pass


def addYearColumn(df):
    # for df_zp only
    print('calculating year duration')
    df['YiPetronas'] = df['Date in PETRONAS']
    df['YiPosition'] = df['Date in Position']
    for i, column in df.iteritems():
        if i == 'YiPetronas' or i == 'YiPosition':
            for index, row in column.iteritems():
                df.loc[index, i] = date_to_duration(row)
    return df


def dropBadColumn(df):
    # for df_zh only
    droplist = ['JG', 'Est.JG', 'EQV JG', 'OLIVE JG', 'Zone', 'NE Area',
                'Vac Date', 'Start Date', 'End Date', 'Vacancy Status']
    df = df.drop(columns=droplist)
    return df


def prepareDataframe(df_zh, df_zp):
    print('preparing dataframe')
    df_zh = dropBadColumn(df_zh)
    df_zp = addYearColumn(df_zp)
    df_zp = df_zp.rename(columns={"Personnel Number": "Staff No",
                                  "Formatted Name of Employee or Applicant": "Staff Name",
                                  "Pos. SKG": "Pos. SKG Zpdev",
                                  "Home SKG": "Home SKG Zpdev"})
    return df_zh, df_zp


def add_spcolumn(df):
    splist = []
    import math
    for i, row in df.iterrows():
        name = row['Staff Name']
        sg = row['SG'] if type(row['SG']) is str else ''
        jg = row['Conso JG'] if type(row['Conso JG']) is str else ''
        possum = row['Pos SUM'] if type(row['Pos SUM']) is str else ''
        retirement = row['Retirement Date'].to_pydatetime().strftime(
            '%d/%m/%Y') if type(row['Retirement Date']) is pd.Timestamp else ''
        age = int(row['Age']) if not math.isnan(row['Age']) else ''
        ppa = row['PPA'] if type(row['PPA']) is str else ''
        sp = '{} ({}, {})\n{}\nRetire: {}, Age: {}\nPPA: {}'.format(
            name, sg, jg, possum, retirement, age, ppa)
        splist.append(sp)
    df['SP'] = splist


print(datetime.now())
opt = input("insert input: \n1. Full DB\n2. Partial DB\nAnswer: ")
if opt == '1':
    zhpla_path = r'D:\Documents\Python\files\ZHPLA_MAY20.xlsx'
    zpdev_path = r'D:\Documents\Python\files\ZPDEV_MAY20.xlsx'
    df_zh = pd.read_excel(zhpla_path, sheet_name='Sheet1', skiprows=3)
    df_zp = pd.read_excel(zpdev_path, sheet_name='Sheet1', skiprows=3)
    filnam = "ta80000.xlsx"
elif opt == '2':
    zhpla_path = r'D:\Documents\Python\files\zh.xlsx'
    zpdev_path = r'D:\Documents\Python\files\zp.xlsx'
    df_zh = pd.read_excel(zhpla_path, sheet_name='Sheet1')
    df_zp = pd.read_excel(zpdev_path, sheet_name='Sheet1')
    filnam = "ta_3000.xlsx"
starttime = datetime.now()
print(df_zh)
print(df_zp)
df_zh, df_zp = prepareDataframe(df_zh, df_zp)
df_ta = merge_database(df_zh, df_zp)
df_ta.to_excel(filnam, index=False, header=True)
endtime = datetime.now()
print('total time elapsed: ', endtime - starttime)
