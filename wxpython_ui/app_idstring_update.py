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
                          pos=wx.DefaultPosition, size=wx.Size(500, 340), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        main_layout = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(
            self, wx.ID_ANY, u"ID Strings", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_title.Wrap(-1)
        self.txt_title.SetFont(
            wx.Font(14, 75, 90, 91, False, "Museo Sans 100"))
        self.txt_title.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        main_layout.Add(self.txt_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        linear_layout_h = wx.BoxSizer(wx.HORIZONTAL)

        b_layout_input = wx.StaticBoxSizer(
            wx.StaticBox(self, wx.ID_ANY, u"Input"), wx.VERTICAL)

        b_layout_selectfile = wx.BoxSizer(wx.HORIZONTAL)

        self.btn_selectfile = wx.Button(b_layout_input.GetStaticBox(
        ), wx.ID_ANY, u"Select file", wx.DefaultPosition, wx.DefaultSize, 0)
        b_layout_selectfile.Add(self.btn_selectfile, 0,
                                wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_filename = wx.StaticText(b_layout_input.GetStaticBox(
        ), wx.ID_ANY, u"filename.xlsx", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_filename.Wrap(-1)
        self.txt_filename.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        b_layout_selectfile.Add(self.txt_filename, 0,
                                wx.ALIGN_CENTER | wx.ALL, 5)

        b_layout_input.Add(b_layout_selectfile, 1, wx.ALL | wx.EXPAND, 5)

        b_layout_specify = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_input.GetStaticBox(), wx.ID_ANY, u"Specify"), wx.VERTICAL)

        b_layout_selection = wx.BoxSizer(wx.HORIZONTAL)

        b_layout_sheet = wx.BoxSizer(wx.VERTICAL)

        self.txt_labelsheet = wx.StaticText(b_layout_specify.GetStaticBox(
        ), wx.ID_ANY, u"Sheet: ", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_labelsheet.Wrap(-1)
        b_layout_sheet.Add(self.txt_labelsheet, 0, wx.ALL, 5)

        cb_selectSheetChoices = []
        self.cb_selectSheet = wx.ComboBox(b_layout_specify.GetStaticBox(
        ), wx.ID_ANY, u"s", wx.DefaultPosition, wx.DefaultSize, cb_selectSheetChoices, wx.CB_READONLY)
        b_layout_sheet.Add(self.cb_selectSheet, 0, wx.ALL, 5)

        b_layout_selection.Add(b_layout_sheet, 1, wx.EXPAND, 5)

        b_layout_column = wx.BoxSizer(wx.VERTICAL)

        self.txt_labelcolumn = wx.StaticText(b_layout_specify.GetStaticBox(
        ), wx.ID_ANY, u"Column: ", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_labelcolumn.Wrap(-1)
        b_layout_column.Add(self.txt_labelcolumn, 0, wx.ALL, 5)

        cb_selectColumnChoices = []
        self.cb_selectColumn = wx.ComboBox(b_layout_specify.GetStaticBox(
        ), wx.ID_ANY, u"Select Column", wx.DefaultPosition, wx.DefaultSize, cb_selectColumnChoices, wx.CB_READONLY)
        b_layout_column.Add(self.cb_selectColumn, 0, wx.ALL, 5)

        b_layout_selection.Add(b_layout_column, 1, wx.EXPAND, 5)

        b_layout_separator = wx.BoxSizer(wx.VERTICAL)

        self.txt_labelseparator = wx.StaticText(b_layout_specify.GetStaticBox(
        ), wx.ID_ANY, u"Separator:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_labelseparator.Wrap(-1)
        b_layout_separator.Add(self.txt_labelseparator, 0, wx.ALL, 5)

        self.txtctrl_separator = wx.TextCtrl(b_layout_specify.GetStaticBox(
        ), wx.ID_ANY, u"", wx.DefaultPosition, wx.DefaultSize, 0)
        b_layout_separator.Add(self.txtctrl_separator, 0, wx.ALL, 5)

        b_layout_selection.Add(b_layout_separator, 1, wx.EXPAND, 5)

        b_layout_specify.Add(b_layout_selection, 1, wx.EXPAND, 5)

        self.txt_message = wx.StaticText(b_layout_specify.GetStaticBox(
        ), wx.ID_ANY, u"", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_message.Wrap(-1)
        self.txt_message.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        b_layout_specify.Add(self.txt_message, 0, wx.ALL, 5)

        b_layout_input.Add(b_layout_specify, 1, wx.EXPAND, 5)

        self.btn_run = wx.Button(b_layout_input.GetStaticBox(
        ), wx.ID_ANY, u"Run!", wx.DefaultPosition, wx.DefaultSize, 0)
        b_layout_input.Add(self.btn_run, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        linear_layout_h.Add(b_layout_input, 1, wx.EXPAND, 5)

        b_layout_output = wx.StaticBoxSizer(
            wx.StaticBox(self, wx.ID_ANY, u"Output"), wx.VERTICAL)

        self.txt_status = wx.StaticText(b_layout_output.GetStaticBox(
        ), wx.ID_ANY, u"Status:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_status.Wrap(-1)
        b_layout_output.Add(self.txt_status, 0,  wx.ALL, 5)

        self.btn_openDir = wx.Button(b_layout_output.GetStaticBox(
        ), wx.ID_ANY, u"Open Folder", wx.DefaultPosition, wx.DefaultSize, 0)
        b_layout_output.Add(self.btn_openDir, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        linear_layout_h.Add(b_layout_output, 1, wx.EXPAND, 5)

        main_layout.Add(linear_layout_h, 1, wx.EXPAND, 5)

        self.SetSizer(main_layout)
        self.Layout()

        self.Centre(wx.BOTH)

        # Variables
        self.mainpath = mainpath
        self.file_link = ''
        self.excelFile = None
        self.currentSheetDF = None

        # Connect Events
        self.btn_selectfile.Bind(wx.EVT_BUTTON, self.selectFile)
        self.btn_run.Bind(wx.EVT_BUTTON, self.runCode)
        self.btn_openDir.Bind(wx.EVT_BUTTON, self.openFolder)
        self.cb_selectSheet.Bind(wx.EVT_COMBOBOX, self.showColumns)

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
            self.file_link = ''.join(paths)
            self.txt_filename.SetLabelText(filename)
            # print("from directory: " + self.file_link)

            self.excelFile = pd.ExcelFile(self.file_link)
            sheetlist = self.excelFile.sheet_names
            self.cb_selectSheet.Clear()
            self.cb_selectSheet.Append(sheetlist)
            self.cb_selectColumn.Clear()
            self.txt_status.SetLabelText("")
            self.txtctrl_separator.SetLabelText("")

        dlg.Destroy()

    def runCode(self, event):
        time, date, time_date = getDate()
        if(self.file_link == None or self.excelFile == None):
            print("No file chosen yet!")
            self.txt_message.SetLabelText("Please choose a file first!")
            self.txt_status.SetLabelText("Status:")
        else:
            sheetname = self.cb_selectSheet.GetValue()
            columnname = self.cb_selectColumn.GetValue()
            if(sheetname == "" or columnname == ""):
                self.txt_message.SetLabelText(
                    "Please select a non empty sheet and specify columns!")
            else:
                self.txt_status.SetLabelText("Status: Running....")
                print("Sheet name: " + sheetname,
                      "\nColumn name: " + columnname)
                df = self.currentSheetDF[[columnname]]
                if(df.empty):
                    self.txt_message.SetLabelText("Sheet is empty")
                    # print("sheet is empty!")
                else:
                    df = df.applymap(str)  # as str
                    # print(df)
                    separator = self.txtctrl_separator.GetValue()
                    if(separator == ""):
                        separator = ","
                    else:
                        separator = self.txtctrl_separator.GetValue()
                    s = df.iloc[:, 0].str.cat(
                        sep=separator)  # get the first column
                    print("Generated on: {}, {}".format(time, date))
                    textfilename = "StaffID ({}).txt".format(time_date)
                    textdir = os.path.join(mainpath, textfilename)
                    id_textfile = open(textdir, "w")
                    id_textfile.write(s)
                    id_textfile.close()
                    self.txt_status.SetLabelText("DONE :D")
                    self.txt_message.SetLabelText("")
                    os.startfile(textdir)

    def openFolder(self, event):
        os.startfile(mainpath)

    def showColumns(self, event):
        self.txt_message.SetLabelText("")
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


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = fr_idToString(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
