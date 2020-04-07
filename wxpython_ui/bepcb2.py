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
###########################################################################
# Class main_frame
###########################################################################
wildcard = "PDF source (*.pdf)|*.pdf|" \
    "All files (*.*)|*.*"


class main_frame (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"BePCB Report", pos=wx.DefaultPosition, size=wx.Size(
            500, 300), style=wx.CAPTION | wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(wx.NORMAL_FONT.GetPointSize(),
                             70, 90, 90, False, "Century Gothic"))
        self.SetBackgroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_ACTIVECAPTION))

        self.menubar = wx.MenuBar(0)
        self.menu_file = wx.Menu()
        self.item_newtask = wx.MenuItem(
            self.menu_file, wx.ID_ANY, u"New Task" + u"\t" + u"Ctrl + N", wx.EmptyString, wx.ITEM_NORMAL)
        self.menu_file.AppendItem(self.item_newtask)

        self.item_exit = wx.MenuItem(
            self.menu_file, wx.ID_ANY, u"Exit" + u"\t" + u"Alt + F4", wx.EmptyString, wx.ITEM_NORMAL)
        self.menu_file.AppendItem(self.item_exit)

        self.menubar.Append(self.menu_file, u"File")

        self.menu_help = wx.Menu()
        self.item_about = wx.MenuItem(
            self.menu_help, wx.ID_ANY, u"About", wx.EmptyString, wx.ITEM_NORMAL)
        self.menu_help.AppendItem(self.item_about)

        self.item_help = wx.MenuItem(
            self.menu_help, wx.ID_ANY, u"Help", wx.EmptyString, wx.ITEM_NORMAL)
        self.menu_help.AppendItem(self.item_help)

        self.menubar.Append(self.menu_help, u"Help")

        self.SetMenuBar(self.menubar)

        boxSizerMain = wx.BoxSizer(wx.VERTICAL)

        self.text_title = wx.StaticText(
            self, wx.ID_ANY, u"Split BePCB Report", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_title.Wrap(-1)
        self.text_title.SetFont(
            wx.Font(9, 74, 90, 90, False, "Century Gothic"))

        boxSizerMain.Add(self.text_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        bsizer_src_kwd = wx.BoxSizer(wx.HORIZONTAL)

        sbsizer_src = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Source file"), wx.VERTICAL)

        bsizer_src = wx.BoxSizer(wx.VERTICAL)

        self.btn_selectfile = wx.Button(sbsizer_src.GetStaticBox(
        ), wx.ID_ANY, u"Select file..", wx.DefaultPosition, wx.DefaultSize, 0)
        bsizer_src.Add(self.btn_selectfile, 0, wx.ALL, 5)

        self.label_filename = wx.StaticText(sbsizer_src.GetStaticBox(
        ), wx.ID_ANY, u"file_name.pdf ?", wx.DefaultPosition, wx.DefaultSize, 0)
        self.label_filename.Wrap(-1)
        bsizer_src.Add(self.label_filename, 0, wx.ALL, 5)

        sbsizer_src.Add(bsizer_src, 1, wx.EXPAND, 5)

        bsizer_src_kwd.Add(sbsizer_src, 1, wx.EXPAND, 5)

        sbsizer_txtfld = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Keyword"), wx.VERTICAL)

        self.textfield_keyword = wx.TextCtrl(sbsizer_txtfld.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        sbsizer_txtfld.Add(self.textfield_keyword, 0, wx.ALL | wx.EXPAND, 5)

        bsizer_src_kwd.Add(sbsizer_txtfld, 1, wx.EXPAND, 5)

        boxSizerMain.Add(bsizer_src_kwd, 1, wx.ALL | wx.EXPAND, 5)

        self.btn_split = wx.Button(
            self, wx.ID_ANY, u"Split!", wx.DefaultPosition, wx.DefaultSize, 0)
        boxSizerMain.Add(self.btn_split, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        sbsizer_output = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Output"), wx.HORIZONTAL)

        bsizer_txt = wx.BoxSizer(wx.VERTICAL)

        self.text_status = wx.StaticText(sbsizer_output.GetStaticBox(
        ), wx.ID_ANY, u"Status: Successful/Unsuccessful", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_status.Wrap(-1)
        bsizer_txt.Add(self.text_status, 0, wx.ALL, 5)

        self.text_reportcount = wx.StaticText(sbsizer_output.GetStaticBox(
        ), wx.ID_ANY, u"Number of report generated: XXX", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_reportcount.Wrap(-1)
        bsizer_txt.Add(self.text_reportcount, 0, wx.ALL, 5)

        sbsizer_output.Add(bsizer_txt, 1, wx.EXPAND, 5)

        bsizer_btnopen = wx.BoxSizer(wx.VERTICAL)

        self.btn_opendir = wx.Button(sbsizer_output.GetStaticBox(
        ), wx.ID_ANY, u"Output Directory...", wx.DefaultPosition, wx.DefaultSize, 0)
        bsizer_btnopen.Add(self.btn_opendir, 0, wx.ALL, 5)

        sbsizer_output.Add(bsizer_btnopen, 1, wx.EXPAND, 5)

        boxSizerMain.Add(sbsizer_output, 1, wx.EXPAND, 5)

        self.SetSizer(boxSizerMain)
        self.Layout()

        self.Centre(wx.BOTH)

        self.currentDirectory = os.getcwd()

        # Connect Events
        self.btn_selectfile.Bind(wx.EVT_BUTTON, self.selectFile)
        self.btn_split.Bind(wx.EVT_BUTTON, self.splitFile)
        self.btn_opendir.Bind(wx.EVT_BUTTON, self.openDirectory)

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def selectFile(self, event):
        dlg = wx.FileDialog(self, message="Choose a file", defaultDir=self.currentDirectory,
                            defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_MULTIPLE | wx.FD_CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            print("You chose the following file(s):")
            pat = dlg.GetFilename()
            self.label_filename.SetLabelText(pat)
            # need to set limit to only 1 file
            # for pa in pat:
            #     self.label_filename.SetLabelText(pa)
            # for path in paths:
            #     print(path)
            #     self.label_filename.SetLabelText(dlg.GetFileNames())
        dlg.Destroy()

    def splitFile(self, event):
        event.Skip()
        # put def of split from bepcb here

    def openDirectory(self, event):
        # insert outputpath hereimport os
        path = "C:/Users"
        path = os.path.realpath(path)
        os.startfile(path)
        # event.Skip()


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = main_frame(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
