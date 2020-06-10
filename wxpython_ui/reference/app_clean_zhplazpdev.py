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
wildcard = "Excel file (*.xlsx)|*.xlsx|" \
    "All files (*.*)|*.*"
###########################################################################
# Class frame_ta
###########################################################################


class frame_ta (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"TTM: Talent Analytics", pos=wx.DefaultPosition, size=wx.Size(
            600, 480), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(9, 74, 90, 90, False, "Century Gothic"))
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        main_layout = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(
            self, wx.ID_ANY, u"Clean ZHPLA/ZPDEV\nTalent Analytics", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTRE)
        self.txt_title.Wrap(-1)
        self.txt_title.SetFont(
            wx.Font(12, 74, 90, 92, False, "Century Gothic"))
        self.txt_title.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        main_layout.Add(self.txt_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        b_layout_h = wx.BoxSizer(wx.VERTICAL)

        b_layout_p1 = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, wx.EmptyString), wx.HORIZONTAL)

        layout_input = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_p1.GetStaticBox(), wx.ID_ANY, u"input"), wx.VERTICAL)

        self.btn_selectzhpla = wx.Button(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"ZHPLA...", wx.DefaultPosition, wx.Size(-1, -1), 0)
        layout_input.Add(self.btn_selectzhpla, 0, wx.ALL, 5)

        self.txt_zhplafile = wx.StaticText(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"zhplafile.xls", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_zhplafile.Wrap(-1)
        layout_input.Add(self.txt_zhplafile, 0, wx.ALL, 5)

        self.btn_selectzpdev = wx.Button(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"ZPDEV...", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_input.Add(self.btn_selectzpdev, 0, wx.ALL, 5)

        self.txt_zpdevfile = wx.StaticText(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"zpdevfile.xls", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_zpdevfile.Wrap(-1)
        layout_input.Add(self.txt_zpdevfile, 0, wx.ALL, 5)

        b_layout_p1.Add(layout_input, 1, wx.EXPAND, 5)

        layout_process = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_p1.GetStaticBox(), wx.ID_ANY, u"process"), wx.VERTICAL)

        layout_choice = wx.BoxSizer(wx.VERTICAL)

        self.rdbtn_clean = wx.RadioButton(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Clean dirty db only", wx.DefaultPosition, wx.DefaultSize, 0)
        self.rdbtn_clean.SetForegroundColour(wx.Colour(0, 177, 169))
        self.rdbtn_clean.SetBackgroundColour(wx.Colour(0, 177, 169))

        layout_choice.Add(self.rdbtn_clean, 0, wx.ALL, 5)

        self.rdbtn_combine = wx.RadioButton(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Combine clean db only", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_choice.Add(self.rdbtn_combine, 0, wx.ALL, 5)

        self.rdbtn_cleancombine = wx.RadioButton(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Clean + combine dirty db", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_choice.Add(self.rdbtn_cleancombine, 0, wx.ALL, 5)

        layout_process.Add(layout_choice, 1, wx.EXPAND, 5)

        layout_button = wx.BoxSizer(wx.VERTICAL)

        self.btn_run = wx.Button(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"text", wx.DefaultPosition, wx.DefaultSize, 0 | wx.NO_BORDER)
        self.btn_run.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        self.btn_run.SetBackgroundColour(wx.Colour(36, 106, 115))

        layout_button.Add(self.btn_run, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_status = wx.StaticText(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Status:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_status.Wrap(-1)
        self.txt_status.SetForegroundColour(wx.Colour(0, 98, 83))

        layout_button.Add(self.txt_status, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        layout_process.Add(layout_button, 1, wx.EXPAND, 5)

        b_layout_p1.Add(layout_process, 1, wx.EXPAND, 5)

        layout_output = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_p1.GetStaticBox(), wx.ID_ANY, u"output"), wx.HORIZONTAL)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        self.btn_openzhpla = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"zhpla", wx.DefaultPosition, wx.Size(60, -1), 0)
        bSizer4.Add(self.btn_openzhpla, 0, wx.ALL, 5)

        self.btn_openzpdev = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"zpdev", wx.DefaultPosition, wx.Size(60, -1), 0)
        bSizer4.Add(self.btn_openzpdev, 0, wx.ALL, 5)

        self.btn_openta = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"talent analytics", wx.DefaultPosition, wx.Size(-1, -1), 0)
        bSizer4.Add(self.btn_openta, 0, wx.ALL, 5)

        layout_output.Add(bSizer4, 1, wx.ALIGN_CENTER | wx.EXPAND, 5)

        self.btn_openoutputdir = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"Open folder", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_output.Add(self.btn_openoutputdir, 0, wx.ALL, 5)

        b_layout_p1.Add(layout_output, 1, wx.EXPAND, 5)

        b_layout_h.Add(b_layout_p1, 1, wx.EXPAND, 5)

        self.txt_message = wx.StaticText(
            self, wx.ID_ANY, u"message here", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_message.Wrap(-1)
        self.txt_message.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        b_layout_h.Add(self.txt_message, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.txt_time_elapsed = wx.StaticText(
            self, wx.ID_ANY, u"time here", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_time_elapsed.Wrap(-1)
        self.txt_time_elapsed.SetFont(
            wx.Font(9, 74, 93, 90, False, "Century Gothic"))
        self.txt_time_elapsed.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_3DLIGHT))

        b_layout_h.Add(self.txt_time_elapsed, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        main_layout.Add(b_layout_h, 1, wx.EXPAND, 5)

        self.SetSizer(main_layout)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.btn_selectzhpla.Bind(wx.EVT_BUTTON, self.selectFile)
        self.btn_selectzpdev.Bind(wx.EVT_BUTTON, self.selectFile)
        self.rdbtn_clean.Bind(wx.EVT_RADIOBUTTON, self.showClean)
        self.rdbtn_combine.Bind(wx.EVT_RADIOBUTTON, self.showCombine)
        self.rdbtn_cleancombine.Bind(wx.EVT_RADIOBUTTON, self.showCleanCombine)
        self.btn_run.Bind(wx.EVT_BUTTON, self.clean_db)
        self.btn_openzhpla.Bind(wx.EVT_BUTTON, self.open_zhpla)
        self.btn_openzpdev.Bind(wx.EVT_BUTTON, self.open_zpdev)
        self.btn_openta.Bind(wx.EVT_BUTTON, self.open_talentAnalytics)
        self.btn_openoutputdir.Bind(wx.EVT_BUTTON, self.OpenOutput)

        self.rdbtn_clean.Hide()
        self.rdbtn_cleancombine.Hide()
        self.rdbtn_combine.Hide()

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def selectFile(self, event):
        event.Skip()

    def showClean(self, event):
        event.Skip()

    def showCombine(self, event):
        event.Skip()

    def showCleanCombine(self, event):
        event.Skip()

    def clean_db(self, event):
        event.Skip()

    def open_zhpla(self, event):
        event.Skip()

    def open_zpdev(self, event):
        event.Skip()

    def open_talentAnalytics(self, event):
        event.Skip()

    def OpenOutput(self, event):
        event.Skip()


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = frame_ta(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
