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
import urllib.request
from urllib.request import Request, urlopen
from urllib.error import URLError
wildcard = "Text file (*.txt)|*.txt|" \
    "All files (*.*)|*.*"
###########################################################################
# Class fr_downloader
###########################################################################


class fr_downloader (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"TTM: TE Batch Downloader",
                          pos=wx.DefaultPosition, size=wx.Size(644, 350), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(
            self, wx.ID_ANY, u"Talent Engine Batch Downloader", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_title.Wrap(-1)
        self.txt_title.SetFont(
            wx.Font(12, 74, 90, 92, False, "Century Gothic"))
        self.txt_title.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        bSizer1.Add(self.txt_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        sbSizer1 = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, wx.EmptyString), wx.HORIZONTAL)

        sbSizer2 = wx.StaticBoxSizer(wx.StaticBox(
            sbSizer1.GetStaticBox(), wx.ID_ANY, u"ID"), wx.VERTICAL)

        # self.txtfld_strings = wx.TextCtrl(sbSizer2.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size(
        #     200, 100), 0 | wx.NO_BORDER | wx.TE_MULTILINE | wx.VSCROLL | wx.BORDER_NONE)
        # sbSizer2.Add(self.txtfld_strings, 0, wx.ALIGN_CENTER |
        #              wx.ALIGN_TOP | wx.ALL, 5)

        self.txtfld_strings = wx.TextCtrl(sbSizer2.GetStaticBox(
        ), wx.ID_ANY, u"input example: 123, 234, 345", wx.DefaultPosition, wx.Size(200, 100), 0 | wx.NO_BORDER | wx.TE_MULTILINE | wx.VSCROLL | wx.BORDER_NONE)
        self.txtfld_strings.SetFont(
            wx.Font(9, 74, 93, 90, False, "Century Gothic"))
        sbSizer2.Add(self.txtfld_strings, 0, wx.ALIGN_CENTER |
                     wx.ALIGN_TOP | wx.ALL, 5)

        self.m_staticText5 = wx.StaticText(sbSizer2.GetStaticBox(
        ), wx.ID_ANY, u"or", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText5.Wrap(-1)
        self.m_staticText5.SetFont(
            wx.Font(8, 74, 93, 90, False, "Century Gothic"))

        sbSizer2.Add(self.m_staticText5, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.btn_selectfile = wx.Button(sbSizer2.GetStaticBox(
        ), wx.ID_ANY, u"Select file...", wx.DefaultPosition, wx.DefaultSize, 0 | wx.NO_BORDER)
        # self.btn_selectfile.SetForegroundColour(wx.Colour(0, 121, 115))
        # self.btn_selectfile.SetBackgroundColour(wx.Colour(216, 216, 216))

        sbSizer2.Add(self.btn_selectfile, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_filepath = wx.StaticText(sbSizer2.GetStaticBox(
        ), wx.ID_ANY, u"file path .txt", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_filepath.Wrap(-1)
        sbSizer2.Add(self.txt_filepath, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        sbSizer1.Add(sbSizer2, 1, wx.EXPAND, 5)

        sbSizer3 = wx.StaticBoxSizer(wx.StaticBox(
            sbSizer1.GetStaticBox(), wx.ID_ANY, u"Selection"), wx.VERTICAL)

        dropdown_filetypeChoices = [u"Select file", u"BePCB", u"CV",
                                    u"DC", u"DW", u"LC", u"Profile Picture"]
        self.cb_filetype = wx.Choice(sbSizer3.GetStaticBox(
        ), wx.ID_ANY, wx.DefaultPosition, wx.Size(180, -1), dropdown_filetypeChoices, 0 | wx.NO_BORDER)
        self.cb_filetype.SetSelection(0)
        # self.dropdown_filetype.SetBackgroundColour(wx.Colour(0, 177, 169))
        self.cb_filetype.SetForegroundColour(wx.Colour(92, 79, 63))
        self.cb_filetype.SetBackgroundColour(wx.Colour(221, 216, 196))

        sbSizer3.Add(self.cb_filetype, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.btn_download = wx.Button(sbSizer3.GetStaticBox(
        ), wx.ID_ANY, u"Download", wx.DefaultPosition, wx.DefaultSize, 0)
        sbSizer3.Add(self.btn_download, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_status = wx.StaticText(sbSizer3.GetStaticBox(
        ), wx.ID_ANY, u"status -....-", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_status.Wrap(-1)
        sbSizer3.Add(self.txt_status, 0,
                     wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 5)

        sbSizer1.Add(sbSizer3, 1, wx.EXPAND, 5)

        sbSizer6 = wx.StaticBoxSizer(wx.StaticBox(
            sbSizer1.GetStaticBox(), wx.ID_ANY, u"Output"), wx.VERTICAL)

        self.btn_openoutputdir = wx.Button(sbSizer6.GetStaticBox(
        ), wx.ID_ANY, u"Open Folder...", wx.Point(-1, -20), wx.DefaultSize, 0)
        sbSizer6.Add(self.btn_openoutputdir, 0, wx.ALIGN_CENTER | wx.TOP, 5)

        sbSizer1.Add(sbSizer6, 1, wx.EXPAND, 5)

        bSizer1.Add(sbSizer1, 1, wx.EXPAND, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # Variables
        self.file_link = ''
        self.mainpath = os.getcwd()
        self.id_list = None
        self.outputFolder = ''

        # Connect Events
        self.btn_selectfile.Bind(wx.EVT_BUTTON, self.selectFile)
        self.btn_download.Bind(wx.EVT_BUTTON, self.download_file)
        self.btn_openoutputdir.Bind(wx.EVT_BUTTON, self.openDir)

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def selectFile(self, event):
        dlg = wx.FileDialog(self, message="Choose a file", defaultDir=self.mainpath,
                            defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            filename = dlg.GetFilename()
            print("You chose the following file(s): %s" % filename)
            self.txt_filepath.SetLabelText(filename)
            self.file_link = ''.join(paths)
            with open(self.file_link, 'r') as file:
                stringtext = file.read().replace('\n', '')

            self.txtfld_strings.Clear()
            self.txtfld_strings.write(stringtext)

            idlist_str = stringtext.split(",")
            idlist_int = []
            for i in idlist_str:
                idlist_int.append(int(i))
            print(idlist_int)
            self.id_list = idlist_int
            # for pa in pat:
            #     self.label_filename.SetLabelText(pa)
            # for path in paths:
            #     print(path)
            #     self.label_filename.SetLabelText(dlg.GetFileNames())
            print(self.file_link)
        dlg.Destroy()

    def download_file(self, event):
        fileindex = self.cb_filetype.GetSelection()
        if(fileindex == 0):
            print("Invalid!")
        else:
            stringtext = self.txtfld_strings.GetValue()
            idlist_str = stringtext.split(",")
            idlist_int = []
            for i in idlist_str:
                idlist_int.append(int(i))
            print(idlist_int)
            self.id_list = idlist_int

            keywordlist = ['', '/bepcb/', '/cv/',
                           '/dc/', '/dw/', '/lc/', '/picture/']
            extensionlist = ['', '.pdf', '.pdf',
                             '.pdf', '.pdf', '.pdf', '.jpg']

            filekeyword = keywordlist[fileindex]
            extension = extensionlist[fileindex]
            fileurl = "https://talentengine.petronas.com/api/resources/staff/"

            self.outputFolder = os.path.join(self.mainpath + filekeyword)
            if not os.path.exists(self.outputFolder):
                os.makedirs(self.outputFolder)
            print(self.outputFolder)
            for id in self.id_list:
                filename = os.path.join(self.outputFolder, str(id) + extension)
                filelink = fileurl + str(id) + filekeyword
                try:
                    urllib.request.urlretrieve(filelink, filename)
                except urllib.error.URLError as e:
                    print(e.code)
                    print(e.read())
                    print(e.reason + " for staff ID: " + str(id))
            print("Done")
            self.txt_status.SetLabelText("Done :D")

    def openDir(self, event):
        os.startfile(self.outputFolder)


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = fr_downloader(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
