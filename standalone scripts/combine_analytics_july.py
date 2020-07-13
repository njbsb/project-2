import pandas as pd
import os
import sys
import time
import datetime


def logDuration(start, end):
    totaltime = int(end - start)
    return str(datetime.timedelta(seconds=totaltime))


def prepareDataframe(zhplapath, zpdevpath):
    starttime = time.time()
    df_zh = pd.read_excel(zhplapath)
    df_zp = pd.read_excel(zpdevpath)

    df_zp = df_zp.rename(columns={'Pos. SKG': 'Pos. SKG Zpdev',
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
                                  })
    endtime = time.time()
    print('prepareDataframe:', logDuration(starttime, endtime))
    return df_zh, df_zp


def merge_database(df_zh, df_zp):
    try:
        starttime = time.time()
        mergeOn = ['Staff No', 'Staff Name']
        df_merge = pd.merge(df_zh, df_zp, on=mergeOn, how='outer')

        # for col in df_merge.columns:
        #     print(col)
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
    finally:
        endtime = time.time()
        print('merge_database:', logDuration(starttime, endtime))


def mainProcess(zhplapath, zpdevpath, outputname):
    starttime = time.time()
    df_zh, df_zp = prepareDataframe(zhplapath, zpdevpath)
    df_merge = merge_database(df_zh, df_zp)
    df_merge.to_excel(outputname, index=None)
    endtime = time.time()
    print('mainProcess:', logDuration(starttime, endtime))
    return df_merge


zhplapath = r'D:\Documents\Python\project-2\zhpla_july.xlsx'
zpdevpath = r'D:\Documents\Python\project-2\combined_zpdevjuly.xlsx'
outputname = 'talent_analytics_month.xlsx'

df_analytics = mainProcess(zhplapath, zpdevpath, outputname)
os.startfile(outputname)
