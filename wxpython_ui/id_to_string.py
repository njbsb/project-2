# -*- coding: utf-8 -*-

###########################################################################
# Python code generated with wxFormBuilder (version Jun 17 2015)
# http://www.wxformbuilder.org/
##
# PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc

###########################################################################
# Class fr_idToString
###########################################################################


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

        self.m_staticText2 = wx.StaticText(b_layout_input.GetStaticBox(
        ), wx.ID_ANY, u"filename.xlsx", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText2.Wrap(-1)
        b_layout_selectfile.Add(self.m_staticText2, 0,
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
        ), wx.ID_ANY, u"unavailable", wx.DefaultPosition, wx.DefaultSize, 0)
        b_layout_separator.Add(self.txtctrl_separator, 0, wx.ALL, 5)

        b_layout_selection.Add(b_layout_separator, 1, wx.EXPAND, 5)

        b_layout_specify.Add(b_layout_selection, 1, wx.EXPAND, 5)

        self.txt_message = wx.StaticText(b_layout_specify.GetStaticBox(
        ), wx.ID_ANY, u"Message:", wx.DefaultPosition, wx.DefaultSize, 0)
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
        b_layout_output.Add(self.txt_status, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.btn_openDir = wx.Button(b_layout_output.GetStaticBox(
        ), wx.ID_ANY, u"Open Folder", wx.DefaultPosition, wx.DefaultSize, 0)
        b_layout_output.Add(self.btn_openDir, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        linear_layout_h.Add(b_layout_output, 1, wx.EXPAND, 5)

        main_layout.Add(linear_layout_h, 1, wx.EXPAND, 5)

        self.SetSizer(main_layout)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.btn_selectfile.Bind(wx.EVT_BUTTON, self.selectFile)
        self.btn_run.Bind(wx.EVT_BUTTON, self.runCode)
        self.btn_openDir.Bind(wx.EVT_BUTTON, self.openFolder)

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def selectFile(self, event):
        event.Skip()

    def runCode(self, event):
        event.Skip()

    def openFolder(self, event):
        event.Skip()


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = fr_idToString(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
