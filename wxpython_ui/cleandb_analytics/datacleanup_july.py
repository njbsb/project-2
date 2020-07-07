import pandas as pd
import time
import os
import sys
import datetime


def logDuration(start, end):
    totaltime = int(end - start)
    return str(datetime.timedelta(seconds=totaltime))


def preProcess(zh='', zp=''):
    datalist = [zh, zp]
    sheetchoice = input('\nSelect db length:\nFull db (1) or partial db (2): ')
    sheet = 'Sheet1' if sheetchoice == '2' else 0
    starttime = time.time()
    rawDFList = []
    for zz in datalist:
        dfRaw = pd.read_excel(
            zz, skiprows=4, sheet_name=sheet) if not zz == '' else None
        rawDFList.append(dfRaw)

    zh_output = 'zhpla_output.xlsx' if not zh == '' else None
    zp_output = 'zpdev_output.xlsx' if not zp == '' else None
    outputname = [zh_output, zp_output]
    endtime = time.time()
    print(rawDFList)
    print('PreProcess: ', logDuration(starttime, endtime))
    return rawDFList, outputname


def cleanOnly():
    pass


def combineOnly(rawDFList):
    pass


zhplapath = r'D:\Documents\Python\project-2\database\input\ZHPLA_June2020_raw.xlsx'
zpdevpath = r'D:\Documents\Python\project-2\database\input\raw_zpdev_june2020.xlsx'

operation = input(
    '\nSelect operation:\nclean (1),combine (2), clean and combine (3): ')


if operation == '1':
    file_option = input(
        '\nSelect files:\nZHPLA (1), ZPDEV (2), ZHPLA and ZPDEV (3): ')
    if file_option == '1':
        rawDFList, outputname = preProcess(zh=zhplapath)
    elif file_option == '2':
        rawDFList, outputname = preProcess(zp=zpdevpath)
    else:
        rawDFList, outputname = preProcess(zh=zhplapath, zp=zpdevpath)

    cleanOnly(rawDFList)
else:
    rawDFList, outputname = preProcess(zh=zhplapath, zp=zpdevpath)
    if operation == '2':
        combineOnly(rawDFList)
    else:
        cleanOnly(rawDFList)
        combineOnly(rawDFList)
