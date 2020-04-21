# -*- coding: utf-8 -*-

###########################################################################
# Python code generated with wxFormBuilder (version Jun 17 2015)
# http://www.wxformbuilder.org/
##
# PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import pandas as pd
import os
wildcard = "Excel File (*.xlsx)|*.xlsx|" \
    "All files (*.*)|*.*"
###########################################################################
# Class MyFrame1
###########################################################################


class MyFrame1 (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title="ID String", pos=wx.DefaultPosition, size=wx.Size(
            500, 200), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_ACTIVECAPTION))

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(
            self, wx.ID_ANY, u"ID String", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_title.Wrap(-1)
        self.txt_title.SetFont(
            wx.Font(wx.NORMAL_FONT.GetPointSize(), 70, 90, 92, False, wx.EmptyString))

        bSizer1.Add(self.txt_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        bSizer2 = wx.BoxSizer(wx.HORIZONTAL)

        sbSizer1 = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Input"), wx.VERTICAL)

        self.btn_selectfile = wx.Button(sbSizer1.GetStaticBox(
        ), wx.ID_ANY, u"Select file", wx.DefaultPosition, wx.DefaultSize, 0)
        sbSizer1.Add(self.btn_selectfile, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_filename = wx.StaticText(sbSizer1.GetStaticBox(
        ), wx.ID_ANY, u"file.xlsx", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_filename.Wrap(-1)
        sbSizer1.Add(self.txt_filename, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        bSizer2.Add(sbSizer1, 1, wx.EXPAND, 5)

        sbSizer3 = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"label"), wx.VERTICAL)

        self.btn_run = wx.Button(sbSizer3.GetStaticBox(
        ), wx.ID_ANY, u"Run", wx.DefaultPosition, wx.DefaultSize, 0)
        sbSizer3.Add(self.btn_run, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 5)

        bSizer2.Add(sbSizer3, 1, wx.EXPAND, 5)

        sbSizer2 = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Output"), wx.VERTICAL)

        self.btn_opendir = wx.Button(sbSizer2.GetStaticBox(
        ), wx.ID_ANY, u"Open Directory", wx.DefaultPosition, wx.DefaultSize, 0)
        sbSizer2.Add(self.btn_opendir, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_status = wx.StaticText(sbSizer2.GetStaticBox(
        ), wx.ID_ANY, u"Status: ", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_status.Wrap(-1)
        sbSizer2.Add(self.txt_status, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        bSizer2.Add(sbSizer2, 1, wx.EXPAND, 5)

        bSizer1.Add(bSizer2, 1, wx.EXPAND, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.btn_selectfile.Bind(wx.EVT_BUTTON, self.selectFile)
        self.btn_run.Bind(wx.EVT_BUTTON, self.runCode)
        self.btn_opendir.Bind(wx.EVT_BUTTON, self.openDir)

        self.label_filename = ""
        self.filepath = ""
        self.currentDirectory = os.getcwd()

        self.btn_opendir.Disable()
        self.btn_run.Disable()

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def selectFile(self, event):
        dlg = wx.FileDialog(self, message="Choose a file", defaultDir=self.currentDirectory,
                            defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()

            filename = dlg.GetFilename()
            print("You chose the following file(s): %s" % filename)
            self.txt_filename.SetLabelText(filename)
            pa = ''.join(paths)
            self.filepath = pa

            self.btn_run.Enable()
            # need to set limit to only 1 file
            # for pa in pat:
            #     self.label_filename.SetLabelText(pa)
            # for path in paths:
            #     print(path)
            #     self.label_filename.SetLabelText(dlg.GetFileNames())
        dlg.Destroy()

    def runCode(self, event):
        idfile = pd.ExcelFile(self.filepath)
        df = pd.read_excel(idfile, usecols='A', sheet_name='Sheet2')
        df = df.applymap(str)
        print(df)
        # s = df['ID'].str.cat(sep=",")
        s = df.iloc[:, 0].str.cat(sep=",")  # get the first column
        # print(s)
        outputfile = "staff_id.txt"
        id_textfile = open(outputfile, "w")
        id_textfile.write(s)
        id_textfile.close()
        os.startfile(os.path.join(os.getcwd(), outputfile))
        self.txt_status.SetLabelText("Status: Success")
        self.btn_opendir.Enable()

    def openDir(self, event):
        outputDir = os.getcwd()
        print(outputDir)
        os.startfile(outputDir)


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = MyFrame1(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
