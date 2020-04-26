# -*- coding: utf-8 -*-

###########################################################################
# Python code generated with wxFormBuilder (version Jun 17 2015)
# http://www.wxformbuilder.org/
##
# PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import os
import pandas as pd
from datetime import datetime
wildcard = "Excel file (*.xlsx)|*.xlsx|" \
    "All files (*.*)|*.*"
mainpath = os.getcwd()
print(mainpath)
###########################################################################
# Class fr_idToString
###########################################################################


def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def getDate():
    dt_tm = str(datetime.now())
    time = dt_tm[11:16]
    timeHR = dt_tm[11:13] + dt_tm[14:16]
    date = dt_tm[8:10] + "-" + dt_tm[5:7] + "-" + dt_tm[0:4]
    dateHR = dt_tm[8:10] + dt_tm[5:7] + dt_tm[0:4]
    time_date = timeHR + "_" + dateHR
    return time, date, time_date


class fr_idToString (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"TTM Tools: ID to Strings",
                          pos=wx.DefaultPosition, size=wx.Size(500, 350), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        # self.m_menubar1 = wx.MenuBar(0)
        # self.menu_newtask = wx.Menu()
        # self.m_menuItem2 = wx.MenuItem(
        #     self.menu_newtask, wx.ID_ANY, u"New Task", wx.EmptyString, wx.ITEM_NORMAL)
        # self.menu_newtask.Append(self.m_menuItem2)

        # self.m_menubar1.Append(self.menu_newtask, u"File")

        # self.SetMenuBar(self.m_menubar1)

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(
            self, wx.ID_ANY, u"ID Strings", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_title.Wrap(-1)
        self.txt_title.SetFont(
            wx.Font(14, 75, 90, 91, False, "Museo Sans 100"))
        self.txt_title.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        bSizer1.Add(self.txt_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        bSizer3 = wx.BoxSizer(wx.HORIZONTAL)

        sbSizer2 = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Input"), wx.VERTICAL)

        bSizer12 = wx.BoxSizer(wx.HORIZONTAL)

        self.btn_selectfile = wx.Button(sbSizer2.GetStaticBox(
        ), wx.ID_ANY, u"Select file", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer12.Add(self.btn_selectfile, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_filename = wx.StaticText(sbSizer2.GetStaticBox(
        ), wx.ID_ANY, u"filename.xlsx", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_filename.Wrap(-1)
        bSizer12.Add(self.txt_filename, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        sbSizer2.Add(bSizer12, 1, wx.EXPAND | wx.ALL, 5)

        sbSizer4 = wx.StaticBoxSizer(wx.StaticBox(
            sbSizer2.GetStaticBox(), wx.ID_ANY, u"Specify"), wx.HORIZONTAL)
        # from here
        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText3 = wx.StaticText(sbSizer4.GetStaticBox(
        ), wx.ID_ANY, u"Sheet: ", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText3.Wrap(-1)
        bSizer4.Add(self.m_staticText3, 0, wx.ALL, 5)

        # self.sheet_name = wx.TextCtrl(sbSizer4.GetStaticBox(
        # ), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        # bSizer4.Add(self.sheet_name, 0, wx.ALL, 5)

        cb_selectSheetChoices = []
        self.cb_selectSheet = wx.ComboBox(sbSizer4.GetStaticBox(
        ), wx.ID_ANY, u"Select sheet", wx.DefaultPosition, wx.DefaultSize, cb_selectSheetChoices, wx.CB_READONLY)
        bSizer4.Add(self.cb_selectSheet, 0, wx.ALL, 5)

        sbSizer4.Add(bSizer4, 1, wx.EXPAND, 5)

        bSizer41 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText31 = wx.StaticText(sbSizer4.GetStaticBox(
        ), wx.ID_ANY, u"Column: ", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText31.Wrap(-1)
        bSizer41.Add(self.m_staticText31, 0, wx.ALL, 5)

        # self.column_name = wx.TextCtrl(sbSizer4.GetStaticBox(
        # ), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        # bSizer41.Add(self.column_name, 0, wx.ALL, 5)

        # self.checkbox_header = wx.CheckBox(sbSizer4.GetStaticBox(
        # ), wx.ID_ANY, u"Table has header", wx.DefaultPosition, wx.DefaultSize, 0)
        # self.checkbox_header.SetForegroundColour(
        #     wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        # self.checkbox_header.SetBackgroundColour(wx.Colour(0, 177, 169))

        # bSizer41.Add(self.checkbox_header, 0, wx.ALL, 5)

        cb_selectColumnChoices = []
        self.cb_selectColumn = wx.ComboBox(sbSizer4.GetStaticBox(
        ), wx.ID_ANY, u"Select Column", wx.DefaultPosition, wx.DefaultSize, cb_selectColumnChoices, wx.CB_READONLY)
        bSizer41.Add(self.cb_selectColumn, 0, wx.ALL, 5)

        sbSizer4.Add(bSizer41, 1, wx.EXPAND, 5)

        bSizer411 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText311 = wx.StaticText(sbSizer4.GetStaticBox(
        ), wx.ID_ANY, u"Separator:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText311.Wrap(-1)
        bSizer411.Add(self.m_staticText311, 0, wx.ALL, 5)

        self.txtctrl_separator = wx.TextCtrl(sbSizer4.GetStaticBox(
        ), wx.ID_ANY, u"not available", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer411.Add(self.txtctrl_separator, 0, wx.ALL, 5)

        sbSizer4.Add(bSizer411, 1, wx.EXPAND, 5)
        # end here
        sbSizer2.Add(sbSizer4, 1, wx.EXPAND, 5)

        self.btn_run = wx.Button(sbSizer2.GetStaticBox(
        ), wx.ID_ANY, u"Run!", wx.DefaultPosition, wx.DefaultSize, 0)
        sbSizer2.Add(self.btn_run, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        bSizer3.Add(sbSizer2, 1, wx.EXPAND, 5)

        sbSizer3 = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Output"), wx.VERTICAL)

        self.txt_status = wx.StaticText(sbSizer3.GetStaticBox(
        ), wx.ID_ANY, u"Status:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_status.Wrap(-1)
        sbSizer3.Add(self.txt_status, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.btn_openDir = wx.Button(sbSizer3.GetStaticBox(
        ), wx.ID_ANY, u"Open Folder", wx.DefaultPosition, wx.DefaultSize, 0)
        sbSizer3.Add(self.btn_openDir, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        bSizer3.Add(sbSizer3, 1, wx.EXPAND, 5)

        bSizer1.Add(bSizer3, 1, wx.EXPAND, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # Variables
        self.mainpath = mainpath
        self.file_link = ''
        self.excelFile = None
        self.currentSheetDF = None
        # self.selectedDF = None

        # Connect Events
        self.btn_selectfile.Bind(wx.EVT_BUTTON, self.selectFile)
        self.btn_run.Bind(wx.EVT_BUTTON, self.runCode)
        self.btn_openDir.Bind(wx.EVT_BUTTON, self.openFolder)
        # self.Bind(wx.EVT_CHECKBOX, self.disableColumnInput)
        self.cb_selectSheet.Bind(wx.EVT_COMBOBOX, self.showColumns)
        # self.cb_selectColumn.Bind(wx.EVT_COMBOBOX, self.show)

        self.txtctrl_separator.Disable()

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def selectFile(self, event):
        self.file_link = ''
        dlg = wx.FileDialog(self, message="Choose a file", defaultDir=self.mainpath,
                            defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            filename = dlg.GetFilename()
            print("You chose the following file(s): %s" % filename)
            self.file_link = ''.join(paths)
            self.txt_filename.SetLabelText(filename)

            self.excelFile = pd.ExcelFile(self.file_link)
            sheetlist = self.excelFile.sheet_names
            self.cb_selectSheet.Clear()
            self.cb_selectSheet.Append(sheetlist)
            self.cb_selectColumn.Clear()
            # for pa in pat:
            #     self.label_filename.SetLabelText(pa)
            # for path in paths:
            #     print(path)
            #     self.label_filename.SetLabelText(dlg.GetFileNames())
            print("from directory: " + self.file_link)
        dlg.Destroy()

    def showColumns(self, event):
        self.cb_selectColumn.Clear()
        sheet = self.cb_selectSheet.GetValue()
        df = pd.read_excel(self.excelFile, sheet_name=sheet)
        self.currentSheetDF = df
        columnlist = list(df.columns)
        # change unnamed to ABC here
        for i, c in enumerate(columnlist):
            if(c[:7] == 'Unnamed'):
                newcol = colnum_string(i+1)
                columnlist[i] = newcol
        self.currentSheetDF.columns = columnlist
        self.cb_selectColumn.Append(columnlist)

    def disableColumnInput(self, event):
        tick = self.checkbox_header.GetValue()
        if(not tick):
            self.column_name.Disable()
        else:
            self.column_name.Enable()

    def runCode(self, event):

        time, date, time_date = getDate()
        self.txt_status.SetLabelText("Status: Running....")
        if(self.file_link == None or self.excelFile == None):
            print("No file chosen yet!")
            self.txt_status.SetLabelText("Status:")
        else:
            sheetname = self.cb_selectSheet.GetValue()
            columnname = self.cb_selectColumn.GetValue()
            if(sheetname == "" or columnname == ""):
                print("shit")
            else:
                print("Sheet name: " + sheetname, "Column name: " + columnname)
                # df = pd.read_excel(file, sheet_name=sheetname)
                df = self.currentSheetDF[[columnname]]
                if(df.empty):
                    print("sheet is empty!")
                else:
                    df = df.applymap(str)  # as str
                    print(df)
                    s = df.iloc[:, 0].str.cat(sep=",")  # get the first column
                    # print(s)
                    print("Generated on: {}, {}".format(time, date))
                    textfilename = "StaffID ({}).txt".format(time_date)
                    textdir = os.path.join(mainpath, textfilename)
                    id_textfile = open(textdir, "w")
                    id_textfile.write(s)
                    id_textfile.close()
                    self.txt_status.SetLabelText("Status: OK")
                    os.startfile(textdir)

    def openFolder(self, event):
        os.startfile(mainpath)


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = fr_idToString(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
