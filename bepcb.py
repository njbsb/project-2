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
wildcard = "PDF source (*.pdf)|*.pdf|" \
    "All files (*.*)|*.*"

###########################################################################
# Class MyFrame2
###########################################################################


class MyFrame2 (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"BePCB Report - Split", pos=wx.DefaultPosition,
                          size=wx.Size(500, 300), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText2 = wx.StaticText(
            self, wx.ID_ANY, u"BePCB Report - Split", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText2.Wrap(-1)
        self.m_staticText2.SetFont(
            wx.Font(wx.NORMAL_FONT.GetPointSize(), 75, 90, 90, False, "Century Gothic"))
        self.m_staticText2.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_CAPTIONTEXT))

        bSizer2.Add(self.m_staticText2, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.m_button5 = wx.Button(
            self, wx.ID_ANY, u"Select file", wx.Point(-1, -1), wx.DefaultSize, 0)
        self.m_button5.SetFont(wx.Font(9, 74, 90, 90, False, "Century Gothic"))

        bSizer2.Add(self.m_button5, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.m_staticText3 = wx.StaticText(
            self, wx.ID_ANY, u"file_name.pdf", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText3.Wrap(-1)
        bSizer2.Add(self.m_staticText3, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.SetSizer(bSizer2)
        self.Layout()
        self.m_menubar1 = wx.MenuBar(0)
        self.m_menu1 = wx.Menu()
        self.m_menubar1.Append(self.m_menu1, u"File")

        self.m_menu2 = wx.Menu()
        self.m_menubar1.Append(self.m_menu2, u"Help")

        self.SetMenuBar(self.m_menubar1)

        self.Centre(wx.BOTH)

        # Connect Events
        self.m_button5.Bind(wx.EVT_BUTTON, self.selectfile)
        self.currentDirectory = os.getcwd()

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def selectfile(self, event):
        dlg = wx.FileDialog(self, message="Choose a file", defaultDir=self.currentDirectory,
                            defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_MULTIPLE | wx.FD_CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            print("You chose the following file(s):")
            for path in paths:
                print(path)
        dlg.Destroy()


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = MyFrame2(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
