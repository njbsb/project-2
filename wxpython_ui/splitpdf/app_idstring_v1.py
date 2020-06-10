# -*- coding: utf-8 -*-

###########################################################################
# Python code generated with wxFormBuilder (version Oct 26 2018)
# http://www.wxformbuilder.org/
##
# PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import pandas as pd
import wx
import wx.xrc
import os
from datetime import datetime
from datetime import date
maindirectory = os.path.dirname(__file__)
# wildcard = "Excel file (*.xlsx)|*.xlsx|" \
#     "All files (*.*)|*.*"
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
    pass


class fr_idToString (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"TTM Tools: ID to Strings",
                          pos=wx.DefaultPosition, size=wx.Size(500, 450), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        main_layout = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(
            self, wx.ID_ANY, u"ID Strings", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_title.Wrap(-1)

        self.txt_title.SetFont(wx.Font(
            16, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "Century Gothic"))
        self.txt_title.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        main_layout.Add(self.txt_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        sz_file = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Select file"), wx.VERTICAL)

        self.filepicker = wx.FilePickerCtrl(sz_file.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, u"Select a file", u"Excel file (*.xlsx)|*.xlsx|"
            "All files (*.*)|*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        sz_file.Add(self.filepicker, 0, wx.ALL | wx.EXPAND, 5)

        self.txt_filename = wx.StaticText(sz_file.GetStaticBox(
        ), wx.ID_ANY, u"filename.xlsx", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_filename.Wrap(-1)

        sz_file.Add(self.txt_filename, 0, wx.ALL, 5)

        main_layout.Add(sz_file, 1, wx.ALL | wx.EXPAND, 5)

        sz_specify = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Specify"), wx.VERTICAL)

        b_layout_selection = wx.BoxSizer(wx.HORIZONTAL)

        b_layout_sheet = wx.BoxSizer(wx.VERTICAL)

        self.txt_labelsheet = wx.StaticText(sz_specify.GetStaticBox(
        ), wx.ID_ANY, u"Sheet: ", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_labelsheet.Wrap(-1)

        b_layout_sheet.Add(self.txt_labelsheet, 0, wx.ALL, 5)

        cb_selectSheetChoices = []
        self.cb_selectSheet = wx.ComboBox(sz_specify.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, cb_selectSheetChoices, wx.CB_READONLY)
        b_layout_sheet.Add(self.cb_selectSheet, 0, wx.ALL, 5)

        b_layout_selection.Add(b_layout_sheet, 1, wx.EXPAND, 5)

        b_layout_column = wx.BoxSizer(wx.VERTICAL)

        self.txt_labelcolumn = wx.StaticText(sz_specify.GetStaticBox(
        ), wx.ID_ANY, u"Column: ", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_labelcolumn.Wrap(-1)

        b_layout_column.Add(self.txt_labelcolumn, 0, wx.ALL, 5)

        cb_selectColumnChoices = []
        self.cb_selectColumn = wx.ComboBox(sz_specify.GetStaticBox(
        ), wx.ID_ANY, u"Select Column", wx.DefaultPosition, wx.DefaultSize, cb_selectColumnChoices, wx.CB_READONLY)
        b_layout_column.Add(self.cb_selectColumn, 0, wx.ALL, 5)

        b_layout_selection.Add(b_layout_column, 1, wx.EXPAND, 5)

        b_layout_separator = wx.BoxSizer(wx.VERTICAL)

        self.txt_labelseparator = wx.StaticText(sz_specify.GetStaticBox(
        ), wx.ID_ANY, u"Separator:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_labelseparator.Wrap(-1)

        b_layout_separator.Add(self.txt_labelseparator, 0, wx.ALL, 5)

        self.txtctrl_separator = wx.TextCtrl(sz_specify.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        b_layout_separator.Add(self.txtctrl_separator, 0, wx.ALL, 5)

        b_layout_selection.Add(b_layout_separator, 1, wx.EXPAND, 5)

        sz_specify.Add(b_layout_selection, 1, wx.EXPAND, 5)

        main_layout.Add(sz_specify, 1, wx.ALL | wx.EXPAND, 5)

        self.btn_run = wx.Button(
            self, wx.ID_ANY, u"Run!", wx.DefaultPosition, wx.DefaultSize, 0)
        main_layout.Add(self.btn_run, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        sz_output = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Output"), wx.HORIZONTAL)

        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        self.btn_openDir = wx.Button(sz_output.GetStaticBox(
        ), wx.ID_ANY, u"Open Folder", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer10.Add(self.btn_openDir, 0,
                     wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 5)

        sz_output.Add(bSizer10, 1, 0, 5)

        main_layout.Add(sz_output, 1, wx.ALIGN_CENTER | wx.ALL | wx.EXPAND, 5)

        self.txt_message = wx.StaticText(
            self, wx.ID_ANY, u"message", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_message.Wrap(-1)

        self.txt_message.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        main_layout.Add(self.txt_message, 0, wx.ALL, 5)

        self.SetSizer(main_layout)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.filepicker.Bind(wx.EVT_FILEPICKER_CHANGED, self.onFileChanged)
        self.btn_run.Bind(wx.EVT_BUTTON, self.runCode)
        self.btn_openDir.Bind(wx.EVT_BUTTON, self.openOutputDir)
        self.cb_selectSheet.Bind(wx.EVT_COMBOBOX, self.showColumns)

        # Enable Disable UI
        self.btn_run.Disable()
        self.btn_openDir.Disable()
        self.cb_selectSheet.Disable()
        self.cb_selectColumn.Disable()
        self.txtctrl_separator.Disable()

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def onFileChanged(self, event):
        filepath = self.filepicker.GetPath()
        if not filepath == '':
            print(filepath)
            filename = os.path.basename(filepath)
            self.txt_filename.SetLabelText(str(filename))
            excelFile = pd.ExcelFile(filepath)
            sheetlist = excelFile.sheet_names

            self.btn_run.Enable()
            self.cb_selectSheet.Enable()
            self.cb_selectColumn.Enable()
            self.txtctrl_separator.Enable()

            self.cb_selectSheet.Clear()
            self.cb_selectSheet.Append(sheetlist)
            self.cb_selectColumn.Clear()
            self.txtctrl_separator.SetLabelText("")
        else:
            self.btn_run.Disable()
            self.cb_selectSheet.Disable()
            self.cb_selectColumn.Disable()
            self.txtctrl_separator.Disable()

    def showColumns(self, event):
        self.txt_message.SetLabelText("")
        self.cb_selectColumn.Clear()
        sheet = self.cb_selectSheet.GetValue()
        df = pd.read_excel(self.filepicker.GetPath(), sheet_name=sheet)
        self.currentSheetDF = df
        columnlist = list(df.columns)
        # change unnamed to ABC here
        for i, c in enumerate(columnlist):
            if(c[:7] == 'Unnamed'):
                newcol = colnum_string(i+1)
                columnlist[i] = newcol
        self.currentSheetDF.columns = columnlist
        self.cb_selectColumn.Append(columnlist)

    def runCode(self, event):
        filepath = self.filepicker.GetPath()
        now = datetime.now()
        datefilename = now.strftime('(%d%m%Y_%H%M)')
        if not filepath == '':
            sheetname = self.cb_selectSheet.GetValue()
            columnname = self.cb_selectColumn.GetValue()
            if not sheetname == '' and not columnname == '':
                try:
                    df = pd.read_excel(filepath, sheet_name=sheetname)[
                        [columnname]]
                    # print(df)
                    if df.empty:
                        self.txt_message.SetLabelText('Sheet is empty')
                    else:
                        try:
                            df = df.dropna(how='any')
                            df = df.astype(int)
                            df = df.applymap(str)
                            separator = self.txtctrl_separator.GetValue()
                            separator = ',' if self.txtctrl_separator.GetValue(
                            ) == '' else separator
                            string = df[columnname].str.cat(sep=separator)
                            try:
                                txtfilename = columnname + datefilename + '.txt'
                                txtfilepath = os.path.join(
                                    maindirectory, txtfilename)
                                with open(txtfilepath, "w") as text_file:
                                    text_file.write(string)
                                os.startfile(txtfilepath)
                                self.txt_message.SetLabelText(
                                    'File executed successfully')
                                self.btn_openDir.Enable()
                            except Exception as e:
                                print(str(e))
                                self.txt_message.SetLabelText(
                                    'Error occured while writing the file')
                        except Exception as e:
                            print(str(e))
                            self.txt_message.SetLabelText(
                                'Error occured while processing the data')
                except Exception as e:
                    print(str(e))
                    self.txt_message.SetLabelText(
                        'Error occured while reading the data')
            else:
                self.txt_message.SetLabelText(
                    'Please select a non empty sheet and specify a column')
        else:
            self.txt_message.SetLabelText('Please select a file first')

    def openOutputDir(self, event):
        os.startfile(maindirectory)


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = fr_idToString(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
