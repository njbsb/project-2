from urllib.error import URLError
from urllib.request import Request, urlopen
import urllib.request
import pandas as pd
import os


def getSheetname(filepath):
    excelFile = pd.ExcelFile(filepath)
    sheetlist = excelFile.sheet_names
    return sheetlist


def getDF_Columns(filepath, sheetname):
    df = pd.read_excel(filepath, sheet_name=sheetname)
    return df, list(df.columns)


def getIDlist(df, columnname):
    idSeries = pd.to_numeric(df[columnname], errors='coerce')
    idlist = idSeries.dropna(how='all').astype(int).tolist()
    return idlist


def downloadFiles(idlist, filetype_index, mainpath):
    resourcepathlist = ['/bepcb/', '/cv/', '/dc/', '/dw/', '/lc/', '/picture/']
    extensionlist = ['.pdf', '.pdf', '.pdf', '.pdf', '.pdf', '.jpg']

    resourcepath = resourcepathlist[filetype_index]
    extension = extensionlist[filetype_index]
    fileurl = "https://talentengine.petronas.com/api/resources/staff/"

    outputFolder = os.path.join(mainpath + resourcepath)
    if not os.path.exists(outputFolder):
        os.makedirs(outputFolder)

    failed_list = []
    successful_list = []
    for idnumber in idlist:
        filename = os.path.join(outputFolder + str(idnumber) + extension)
        filelink = fileurl + str(idnumber) + resourcepath
        try:
            urllib.request.urlretrieve(filelink, filename)
            print('Downloading from', filelink)
            successful_list.append(filename)
        except urllib.error.URLError as e:
            print(e.code)
            print(e.read())
            print(e.reason + " for staff ID: " + str(idnumber))
            failed_list.append(idnumber)
    return successful_list, failed_list, outputFolder
