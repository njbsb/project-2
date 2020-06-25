import pandas as pd
import os
import sys
from datetime import date, datetime
from dateutil.relativedelta import relativedelta


def date_to_duration(date):
    dip = date.date()
    rdelta = relativedelta(date.today(), dip)
    dateYM = str(rdelta.years) + "y" + str(rdelta.months) + "m"
    return dateYM


def merge_database(df_zh, df_zp):
    try:
        print('merging databases')
        mergeOn = ['Staff No', 'Staff Name']
        df_merge = pd.merge(df_zh, df_zp, on=mergeOn, how='outer')
        # print(1)
        compare_list = [['SG_x', 'SG_y'], ['Lvl_x', 'Lvl_y'], ['Position_x', 'Position_y'], ['Section_x', 'Section_y'], [
            'Department_x', 'Department_y'], ['Division_x', 'Division_y'], ['Sector_x', 'Sector_y'], ['Gender_x', 'Gender_y']]
        for cl in compare_list:
            df_merge[cl[0]] = df_merge.apply(
                lambda row: row[cl[1]] if pd.isnull(row[cl[0]]) else row[cl[0]], axis=1)
            xcolumn = cl[0]
            if(xcolumn[-2:] == '_x'):
                df_merge.rename(columns={xcolumn: xcolumn[:-2]}, inplace=True)
        # print(2)
        JG_cl = ['Conso JG', 'JG']
        df_merge[JG_cl[0]] = df_merge.apply(lambda row: row[JG_cl[1]] if pd.isnull(
            row[JG_cl[0]]) or row[JG_cl[0]] == 'No JG' else row[JG_cl[0]], axis=1)
        # df_merge['c'] = df_merge.apply(
        #     lambda row: row['a']*row['b'] if pd.isnull(row['c']) else row['c'], axis=1)
        # print(3)
        drop_column = [cl[1] for cl in compare_list]
        drop_column.append(JG_cl[1])
        for drop in drop_column:
            df_merge = df_merge.drop(drop, axis=1)
        # print(4)
        arrangedcol = ['Staff No', 'Staff Name', 'Gender', 'Race', 'Age', 'SG', 'Lvl', 'Tier', 'Conso JG', 'Pos ID', 'Pos SUM', 'Superior Pos ID', 'Superior ID', 'Superior Name', 'SG Since', 'SG (YM)', 'YiPETRONAS', 'YiPosition',
                       'PPA-19', 'PPA-18', 'PPA-17', 'PPA-16', 'PPA-15', 'PPA', 'Date in PETRONAS', 'Date in Position', 'Date in Department', 'Date in Division', 'Date in Company', 'Date of Birth', 'Country of Birth', 'Country of Birth Desc',
                       'State', 'State Desc', 'Birthplace', 'Nationality', 'Nationality Desc', 'Position', 'Sec ID', 'Section', 'Dept ID', 'Department', 'Division ID', 'Division', 'Sector ID', 'Sector', 'Comp. ID Position', 'Comp. Position', 'Buss ID', 'Business',
                       'C.Center', 'Home SKG', 'Pos.SKG', 'Pos. SKG Zpdev', 'Home SKG Zpdev', 'JobID(C)', 'Level', 'Loc ID', 'Location Text', 'Work Center', 'Email Address', 'Task Type', 'Task Type Desc', 'Retirement Date']
        dfm_col = list(df_merge.columns)
        check = all(item in dfm_col for item in arrangedcol)
        print('Check:', check)
        for item in arrangedcol:
            if item not in dfm_col:
                print(item)

        # set1 = set(dfm_col)
        # set2 = set(arrangedcol)
        # print('difference', str(set2.difference(set1)))
        try:

            df_merge = df_merge[arrangedcol]
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(
                exc_tb.tb_frame.f_code.co_filename)[1]
            print('Error')
            print(exc_type, fname, exc_tb.tb_lineno)
        df_merge = df_merge.sort_values(by=['Pos ID'], ascending=True)
        # add_spcolumn(df_merge)
        print('writing to excel')
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(
            exc_tb.tb_frame.f_code.co_filename)[1]
        print('Error')
        print(exc_type, fname, exc_tb.tb_lineno)
    return df_merge


def addYearColumn(df):
    # for df_zp only
    print('adding year column')
    dateIn = ['Date in PETRONAS', 'Date in Position']
    yearIn = ['YiPETRONAS', 'YiPosition']
    try:
        for di, yi in zip(dateIn, yearIn):
            yi_list = []
            for eachrow in df[di]:
                #  or not type(eachrow) is float
                if type(eachrow) is pd.Timestamp:
                    yi_list.append(date_to_duration(eachrow))
                else:
                    yi_list.append(eachrow)
            df[yi] = yi_list
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(
            exc_tb.tb_frame.f_code.co_filename)[1]
        print('Error')
        print(exc_type, fname, exc_tb.tb_lineno)
    print('done add YI columns')
    return df


def dropBadColumn(df):
    # for df_zh only
    droplist = ['JG', 'Est.JG', 'EQV JG', 'OLIVE JG', 'Zone', 'NE Area',
                'Vac Date', 'Start Date', 'End Date', 'Vacancy Status']
    df = df.drop(columns=droplist)
    print('done drop bad column from zh')
    return df


def prepareDataframe(df_zh, df_zp):
    print('preparing dataframe')
    df_zh = dropBadColumn(df_zh)
    df_zp = addYearColumn(df_zp)
    df_zp = df_zp.rename(columns={'Pos. SKG': 'Pos. SKG Zpdev',
                                  'Home SKG': 'Home SKG Zpdev',
                                  'Staff Number': 'Staff No',
                                  'Date of Task': 'Retirement Date'
                                  })
    print('renaming completed')
    return df_zh, df_zp


def add_spcolumn(df):
    splist = []
    import math
    try:
        for i, row in df.iterrows():
            name = row['Staff Name']
            sg = row['SG'] if type(row['SG']) is str else ''
            jg = row['Conso JG'] if type(row['Conso JG']) is str else ''
            possum = row['Pos SUM'] if type(row['Pos SUM']) is str else ''

            retirement = row['Retirement Date'].to_pydatetime().strftime(
                '%d/%m/%Y') if type(row['Retirement Date']) is pd.Timestamp else ''
            # retirement = row['Retirement Date'] if type(
            #     row['Retirement Date']) is str else ''
            print('retirement success')
            age = int(row['Age']) if not math.isnan(row['Age']) else ''
            ppa = row['PPA'] if type(row['PPA']) is str else ''
            sp = '{} ({}, {})\n{}\nRetire: {}, Age: {}\nPPA: {}'.format(
                name, sg, jg, possum, retirement, age, ppa)
            splist.append(sp)
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(
            exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
    df['SP'] = splist


# print(datetime.now())
# opt = input("insert input: \n1. Full DB\n2. Partial DB\nAnswer: ")
# if opt == '1':
#     zhpla_path = r'D:\Documents\Python\files\ZHPLA_MAY20.xlsx'
#     zpdev_path = r'D:\Documents\Python\files\ZPDEV_MAY20.xlsx'
#     df_zh = pd.read_excel(zhpla_path, sheet_name='Sheet1', skiprows=3)
#     df_zp = pd.read_excel(zpdev_path, sheet_name='Sheet1', skiprows=3)
#     filnam = "ta80000.xlsx"
# elif opt == '2':
#     zhpla_path = r'D:\Documents\Python\files\zh.xlsx'
#     zpdev_path = r'D:\Documents\Python\files\zp.xlsx'
#     df_zh = pd.read_excel(zhpla_path, sheet_name='Sheet1')
#     df_zp = pd.read_excel(zpdev_path, sheet_name='Sheet1')
#     filnam = "ta_3000.xlsx"
# starttime = datetime.now()
# df_zh, df_zp = prepareDataframe(df_zh, df_zp)
# df_ta = merge_database(df_zh, df_zp)
# df_ta.to_excel(filnam, index=False, header=True)
# endtime = datetime.now()
# print('total time elapsed: ', endtime - starttime)
