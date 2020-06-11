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
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"TTM: TE Downloader",
                          pos=wx.DefaultPosition, size=wx.Size(644, 350), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)
        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        main_layout = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(
            self, wx.ID_ANY, u"Talent Engine Downloader", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_title.Wrap(-1)
        self.txt_title.SetFont(
            wx.Font(12, 74, 90, 92, False, "Century Gothic"))
        self.txt_title.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        main_layout.Add(self.txt_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        b_layout_h = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, wx.EmptyString), wx.HORIZONTAL)

        b_layout_ID = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_h.GetStaticBox(), wx.ID_ANY, u"ID"), wx.VERTICAL)

        self.txtfld_strings = wx.TextCtrl(b_layout_ID.GetStaticBox(
        ), wx.ID_ANY, u"input example: 123, 234, 345", wx.DefaultPosition, wx.Size(200, 100), 0 | wx.NO_BORDER | wx.TE_MULTILINE | wx.VSCROLL | wx.BORDER_NONE)
        self.txtfld_strings.SetFont(wx.Font(9, 74, 93, 90, False, "Arial"))

        b_layout_ID.Add(self.txtfld_strings, 0,
                        wx.ALIGN_CENTER | wx.ALIGN_TOP | wx.ALL, 5)

        self.m_staticText5 = wx.StaticText(b_layout_ID.GetStaticBox(
        ), wx.ID_ANY, u"or", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText5.Wrap(-1)
        self.m_staticText5.SetFont(
            wx.Font(8, 74, 93, 90, False, "Century Gothic"))

        b_layout_ID.Add(self.m_staticText5, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.btn_selectfile = wx.Button(b_layout_ID.GetStaticBox(
        ), wx.ID_ANY, u"Select txt file...", wx.DefaultPosition, wx.DefaultSize, 0 | wx.NO_BORDER)
        self.btn_selectfile.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        self.btn_selectfile.SetBackgroundColour(wx.Colour(36, 106, 115))

        b_layout_ID.Add(self.btn_selectfile, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_filepath = wx.StaticText(b_layout_ID.GetStaticBox(
        ), wx.ID_ANY, u"file path .txt", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_filepath.Wrap(-1)
        b_layout_ID.Add(self.txt_filepath, 0, wx.ALL, 5)

        self.txt_countID = wx.StaticText(b_layout_ID.GetStaticBox(
        ), wx.ID_ANY, u"", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_countID.Wrap(-1)
        self.txt_countID.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        b_layout_ID.Add(self.txt_countID, 0, wx.ALL, 5)

        b_layout_h.Add(b_layout_ID, 1, wx.EXPAND, 5)

        b_layout_selection = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_h.GetStaticBox(), wx.ID_ANY, u"Selection"), wx.VERTICAL)

        dropdown_filetypeChoices = [u"Select file",
                                    u"BePCB", u"CV", u"DC", u"DW", u"LC", u"Profile Picture"]
        self.dropdown_filetype = wx.Choice(b_layout_selection.GetStaticBox(
        ), wx.ID_ANY, wx.DefaultPosition, wx.Size(180, -1), dropdown_filetypeChoices, 0 | wx.NO_BORDER)
        self.dropdown_filetype.SetSelection(0)
        self.dropdown_filetype.SetForegroundColour(wx.Colour(92, 79, 63))
        self.dropdown_filetype.SetBackgroundColour(wx.Colour(221, 216, 196))

        b_layout_selection.Add(self.dropdown_filetype, 0,
                               wx.ALIGN_CENTER | wx.ALL, 5)

        self.btn_download = wx.Button(b_layout_selection.GetStaticBox(
        ), wx.ID_ANY, u"Download", wx.DefaultPosition, wx.DefaultSize, 0 | wx.NO_BORDER)
        self.btn_download.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        self.btn_download.SetBackgroundColour(wx.Colour(36, 106, 115))

        b_layout_selection.Add(self.btn_download, 0,
                               wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_status = wx.StaticText(b_layout_selection.GetStaticBox(
        ), wx.ID_ANY, u"", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_status.Wrap(-1)
        self.txt_status.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        b_layout_selection.Add(self.txt_status, 0, wx.ALL, 5)

        self.txt_countRemaining = wx.StaticText(b_layout_selection.GetStaticBox(
        ), wx.ID_ANY, u"", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_countRemaining.Wrap(-1)
        self.txt_countRemaining.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        b_layout_selection.Add(self.txt_countRemaining, 0, wx.ALL, 5)

        b_layout_h.Add(b_layout_selection, 1, wx.EXPAND, 5)

        self.txtfld_failed = wx.TextCtrl(b_layout_selection.GetStaticBox(
        ), wx.ID_ANY, u"", wx.DefaultPosition, wx.Size(200, 100), 0 | wx.NO_BORDER | wx.TE_MULTILINE | wx.VSCROLL | wx.BORDER_NONE)
        self.txtfld_failed.SetForegroundColour(wx.Colour(36, 106, 115))
        self.txtfld_failed.SetBackgroundColour(wx.Colour(0, 177, 169))
        self.txtfld_failed.Enable(True)

        b_layout_selection.Add(self.txtfld_failed, 0, wx.ALL, 5)

        b_layout_output = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_h.GetStaticBox(), wx.ID_ANY, u"Output"), wx.VERTICAL)

        self.txt_statusFinal = wx.StaticText(b_layout_output.GetStaticBox(
        ), wx.ID_ANY, u"", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_statusFinal.Wrap(-1)
        self.txt_statusFinal.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        b_layout_output.Add(self.txt_statusFinal, 0,
                            wx.ALIGN_CENTER | wx.ALL, 5)

        self.btn_openoutputdir = wx.Button(b_layout_output.GetStaticBox(
        ), wx.ID_ANY, u"Open Folder...", wx.Point(-1, -20), wx.DefaultSize, 0)
        b_layout_output.Add(self.btn_openoutputdir, 0,
                            wx.ALIGN_CENTER | wx.TOP, 5)

        b_layout_h.Add(b_layout_output, 1, wx.EXPAND, 5)

        main_layout.Add(b_layout_h, 1, wx.EXPAND, 5)

        self.SetSizer(main_layout)
        self.Layout()

        self.Centre(wx.BOTH)

        # Variables
        self.file_link = ''
        self.mainpath = os.path.dirname(__file__)
        self.id_list = None
        self.outputFolder = self.mainpath

        # Connect Events
        self.btn_selectfile.Bind(wx.EVT_BUTTON, self.selectFile)
        self.btn_download.Bind(wx.EVT_BUTTON, self.download_file)
        self.btn_openoutputdir.Bind(wx.EVT_BUTTON, self.openDir)
        self.txtfld_strings.Bind(wx.EVT_TEXT, self.update_ID_count)
        self.txtfld_strings.Bind(wx.EVT_CHAR, self.OnChar)
        self.txt_filepath.SetLabelText("No file chosen")

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
            self.id_list = idlist_int
            # print(self.file_link)
        dlg.Destroy()

    def download_file(self, event):
        fileindex = self.dropdown_filetype.GetSelection()
        if(fileindex == 0):
            print("Invalid!")
        else:
            stringtext = self.txtfld_strings.GetValue()
            stringtext = stringtext.replace(" ", "")
            if(stringtext[-1] == ','):
                stringtext = stringtext[:-1]
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
            failed_list = []
            for id in self.id_list:
                filename = os.path.join(self.outputFolder, str(id) + extension)
                filelink = fileurl + str(id) + filekeyword
                remaining = len(self.id_list) - self.id_list.index(id)
                self.txt_status.SetLabelText("Downloading: %s" % str(id))
                self.txt_countRemaining.SetLabelText(str(remaining))
                try:
                    urllib.request.urlretrieve(filelink, filename)
                except urllib.error.URLError as e:
                    print(e.code)
                    print(e.read())
                    print(e.reason + " for staff ID: " + str(id))
                    failed_list.append(id)
            self.txt_status.SetLabelText("Status: Done")
            if not failed_list:
                self.txt_countRemaining.SetLabelText("Download successful")
                self.txtfld_failed.SetLabelText('')
                os.startfile(self.outputFolder)
            else:
                self.txt_countRemaining.SetLabelText(
                    "Failed: {}".format(len(failed_list)))
                self.txtfld_failed.SetLabelText(str(failed_list))

    def openDir(self, event):
        os.startfile(self.outputFolder)

    def OnChar(self, event):
        key = event.GetKeyCode()
        acceptable_characters = "1234567890,"
        acceptable_key = [wx.WXK_BACK, wx.WXK_LEFT,
                          wx.WXK_RIGHT, wx.WXK_UP, wx.WXK_DOWN]
        if chr(key) in acceptable_characters or key in acceptable_key:
            event.Skip()
            return
        else:
            return False

    def update_ID_count(self, event):
        ids = self.txtfld_strings.GetValue()
        if(ids == ''):
            self.txt_countID.SetLabelText('')
        else:
            if(ids[-1] == ','):
                ids = ids[:-1]
            countID = ids.count(',') + 1
            if(countID != None or countID != 0):
                self.txt_countID.SetLabelText(
                    "Number of ID: %s" % str(countID))


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = fr_downloader(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
