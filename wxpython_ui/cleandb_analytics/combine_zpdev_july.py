import pandas as pd
import os
import zpdevcleanup_july as zp


def mainProcess(zpdevmain_path, zpdevretire_path, outputname):
    df_main = pd.read_excel(zpdevmain_path)
    df_retire = pd.read_excel(zpdevretire_path)

    dflist = [df_main, df_retire]
    for i, df in enumerate(dflist):
        df_group = zp.getDFGroup(df)
        cleandf = zp.getCleanDF(df_group)
        dflist[i] = cleandf

    df_main, df_retire = dflist
    combined_zpdev = pd.merge(df_main, df_retire[[
        'Personnel Number', 'Task Type', 'Task Type Desc', 'Date of Task']], on='Personnel Number', how='left')
    combined_zpdev = zp.addcolumn_ppa5yrs(combined_zpdev)
    combined_zpdev = zp.addcolumn_SGym(combined_zpdev)
    combined_zpdev = zp.addcolumn_age(combined_zpdev)
    combined_zpdev = zp.addcolumn_yearIn(combined_zpdev)
    combined_zpdev = zp.arrange_column(combined_zpdev)

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
