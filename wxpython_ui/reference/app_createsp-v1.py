# -*- coding: utf-8 -*-

###########################################################################
# Python code generated with wxFormBuilder (version Oct 26 2018)
# http://www.wxformbuilder.org/
##
# PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import pandas as pd
import os
wildcard = "Excel file (*.xlsx)|*.xlsx|" \
    "All files (*.*)|*.*"
mainpath = os.getcwd()
###########################################################################
# Class mainframe
###########################################################################


class mainframe (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"Succession Planning Pack Generator",
                          pos=wx.DefaultPosition, size=wx.Size(500, 300), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        mainsizer = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(
            self, wx.ID_ANY, u"SP PACK Generator", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_title.Wrap(-1)

        self.txt_title.SetFont(wx.Font(
            11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "Century Gothic"))
        self.txt_title.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        mainsizer.Add(self.txt_title, 0,
                      wx.ALIGN_CENTER_HORIZONTAL | wx.BOTTOM | wx.TOP, 5)

        content_sz = wx.BoxSizer(wx.HORIZONTAL)

        input_sz = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Input"), wx.VERTICAL)

        self.btn_selectfile = wx.Button(input_sz.GetStaticBox(
        ), wx.ID_ANY, u"Select sp excel...", wx.DefaultPosition, wx.DefaultSize, 0)
        input_sz.Add(self.btn_selectfile, 0, wx.ALL, 5)

        self.txt_filelabel = wx.StaticText(input_sz.GetStaticBox(
        ), wx.ID_ANY, u"*filelabel*", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_filelabel.Wrap(-1)

        self.txt_filelabel.SetFont(wx.Font(
            9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))

        input_sz.Add(self.txt_filelabel, 0, wx.ALL, 5)

        self.btn_mediafolder = wx.Button(input_sz.GetStaticBox(
        ), wx.ID_ANY, u"Set Media Folder", wx.DefaultPosition, wx.DefaultSize, 0)
        input_sz.Add(self.btn_mediafolder, 0, wx.ALL, 5)

        self.txt_mediapath = wx.StaticText(input_sz.GetStaticBox(
        ), wx.ID_ANY, u"*media path*", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_mediapath.Wrap(-1)

        self.txt_mediapath.SetFont(wx.Font(
            9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))

        input_sz.Add(self.txt_mediapath, 0, wx.ALL, 5)

        content_sz.Add(input_sz, 1, wx.EXPAND, 5)

        process_sz = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Process"), wx.VERTICAL)

        self.btn_generate = wx.Button(process_sz.GetStaticBox(
        ), wx.ID_ANY, u"Generate SP", wx.DefaultPosition, wx.DefaultSize, 0)
        process_sz.Add(self.btn_generate, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        content_sz.Add(process_sz, 1, wx.ALIGN_CENTER |
                       wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        output_sz = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Output"), wx.VERTICAL)

        self.btn_openfolder = wx.Button(output_sz.GetStaticBox(
        ), wx.ID_ANY, u"Open folder", wx.DefaultPosition, wx.DefaultSize, 0)
        output_sz.Add(self.btn_openfolder, 0, wx.ALL, 5)

        content_sz.Add(output_sz, 1, wx.EXPAND, 5)

        mainsizer.Add(content_sz, 1, wx.EXPAND |
                      wx.LEFT | wx.RIGHT | wx.TOP, 5)

        self.txt_message = wx.StaticText(
            self, wx.ID_ANY, u"message", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_message.Wrap(-1)

        self.txt_message.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_3DLIGHT))

        mainsizer.Add(self.txt_message, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.SetSizer(mainsizer)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.btn_selectfile.Bind(wx.EVT_BUTTON, self.selectFile)
        self.btn_mediafolder.Bind(wx.EVT_BUTTON, self.selectMediaFolder)
        self.btn_generate.Bind(wx.EVT_BUTTON, self.runSP)
        self.btn_openfolder.Bind(wx.EVT_BUTTON, self.openOutputFolder)

        self.mainpath = mainpath

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def selectFile(self, event):
        dlg = wx.FileDialog(self, message="Choose a file", defaultDir=self.mainpath,
                            defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            filename = dlg.GetFilename()
            # print("You chose the following file(s): %s" % filename)
            # self.txt_filename.SetLabelText(filename)
            # pa = ''.join(paths)
            # self.filepath = pa

            # self.btn_run.Enable()
            # need to set limit to only 1 file
            # for pa in pat:
            #     self.label_filename.SetLabelText(pa)
            # for path in paths:
            #     print(path)
            #     self.label_filename.SetLabelText(dlg.GetFileNames())
        dlg.Destroy()

    def selectMediaFolder(self, event):
        dlg = wx.FileDialog(self, message="Choose media folder path", defaultDir=self.mainpath,
                            defaultFile="", style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            filename = dlg.GetFilename()
            # print("You chose the following file(s): %s" % filename)
            # self.txt_filename.SetLabelText(filename)
            # pa = ''.join(paths)
            # self.filepath = pa

            # self.btn_run.Enable()
            # need to set limit to only 1 file
            # for pa in pat:
            #     self.label_filename.SetLabelText(pa)
            # for path in paths:
            #     print(path)
            #     self.label_filename.SetLabelText(dlg.GetFileNames())
        dlg.Destroy()

    def runSP(self, event):
        event.Skip()

    def openOutputFolder(self, event):
        event.Skip()


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = mainframe(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
