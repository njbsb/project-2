# -*- coding: utf-8 -*-

###########################################################################
# Python code generated with wxFormBuilder (version Oct 26 2018)
# http://www.wxformbuilder.org/
##
# PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import os
import module_analytics_july as dc
###########################################################################
# Class AnalyticsFrame
###########################################################################


class AnalyticsFrame (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"TTM: Talent Analytics", pos=wx.DefaultPosition, size=wx.Size(
            600, 480), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        main_layout = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(self, wx.ID_ANY, u"Top Talent Analytics",
                                       wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL)
        self.txt_title.Wrap(-1)

        self.txt_title.SetFont(wx.Font(
            12, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "Century Gothic"))
        self.txt_title.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        main_layout.Add(self.txt_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        b_layout_h = wx.BoxSizer(wx.VERTICAL)

        b_layout_p1 = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, wx.EmptyString), wx.HORIZONTAL)

        layout_input = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_p1.GetStaticBox(), wx.ID_ANY, u"input"), wx.VERTICAL)

        self.zhpla_picker = wx.FilePickerCtrl(layout_input.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        layout_input.Add(self.zhpla_picker, 0, wx.ALL, 5)

        self.txt_zhplafile = wx.StaticText(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"ZHPLA file (.xlsx)", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_zhplafile.Wrap(-1)

        layout_input.Add(self.txt_zhplafile, 0, wx.ALL, 5)

        self.zpdev_picker = wx.FilePickerCtrl(layout_input.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        layout_input.Add(self.zpdev_picker, 0, wx.ALL, 5)

        self.txt_zpdevfile = wx.StaticText(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"ZPDEV main (xlsx)", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_zpdevfile.Wrap(-1)

        layout_input.Add(self.txt_zpdevfile, 0, wx.ALL, 5)

        self.zpdevretire_picker = wx.FilePickerCtrl(layout_input.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        layout_input.Add(self.zpdevretire_picker, 0, wx.ALL, 5)

        self.txt_zpdevretirefile = wx.StaticText(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"ZPDEV retire (xlsx)", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_zpdevretirefile.Wrap(-1)

        layout_input.Add(self.txt_zpdevretirefile, 0, wx.ALL, 5)

        b_layout_p1.Add(layout_input, 1, wx.EXPAND, 5)

        layout_process = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_p1.GetStaticBox(), wx.ID_ANY, u"select process"), wx.VERTICAL)

        layout_choice = wx.BoxSizer(wx.VERTICAL)

        self.rdbtn_clean = wx.RadioButton(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Clean dirty db only", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_choice.Add(self.rdbtn_clean, 0, wx.ALL, 5)

        self.rdbtn_combine = wx.RadioButton(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Combine clean db only", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_choice.Add(self.rdbtn_combine, 0, wx.ALL, 5)

        self.rdbtn_cleancombine = wx.RadioButton(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Clean + combine dirty db", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_choice.Add(self.rdbtn_cleancombine, 0, wx.ALL, 5)

        self.btn_autoopen = wx.CheckBox(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Open when done", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_choice.Add(self.btn_autoopen, 0, wx.ALL, 5)

        layout_process.Add(layout_choice, 1, wx.EXPAND, 5)

        layout_button = wx.BoxSizer(wx.VERTICAL)

        self.btn_run = wx.Button(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Clean / Combine", wx.DefaultPosition, wx.Size(120, -1), 0)
        self.btn_run.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWTEXT))
        self.btn_run.SetBackgroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_3DLIGHT))

        layout_button.Add(self.btn_run, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        # self.txt_status = wx.StaticText(layout_process.GetStaticBox(
        # ), wx.ID_ANY, u"Status:", wx.DefaultPosition, wx.DefaultSize, 0)
        # self.txt_status.Wrap(-1)

        # self.txt_status.SetForegroundColour(wx.Colour(0, 98, 83))

        # layout_button.Add(self.txt_status, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        layout_process.Add(layout_button, 1, wx.EXPAND, 5)

        b_layout_p1.Add(layout_process, 1, wx.EXPAND, 5)

        layout_output = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_p1.GetStaticBox(), wx.ID_ANY, u"output"), wx.VERTICAL)

        self.btn_openzhpla = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"zhpla", wx.DefaultPosition, wx.Size(60, -1), 0)
        layout_output.Add(self.btn_openzhpla, 0, wx.ALL, 5)

        self.btn_openzpdev = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"zpdev", wx.DefaultPosition, wx.Size(60, -1), 0)
        layout_output.Add(self.btn_openzpdev, 0, wx.ALL, 5)

        self.btn_openta = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"talent analytics", wx.DefaultPosition, wx.Size(-1, -1), 0)
        layout_output.Add(self.btn_openta, 0, wx.ALL, 5)

        self.btn_openoutputdir = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"Open folder", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_output.Add(self.btn_openoutputdir, 0, wx.ALL, 5)

        b_layout_p1.Add(layout_output, 1, wx.EXPAND, 5)

        b_layout_h.Add(b_layout_p1, 1, wx.EXPAND, 5)

        self.txt_message = wx.StaticText(
            self, wx.ID_ANY, u"Welcome to TTAnalytics", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_message.Wrap(-1)

        self.txt_message.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        b_layout_h.Add(self.txt_message, 0, wx.ALL, 5)

        # self.txt_time_elapsed = wx.StaticText(
        #     self, wx.ID_ANY, u"time here", wx.DefaultPosition, wx.DefaultSize, 0)
        # self.txt_time_elapsed.Wrap(-1)

        # self.txt_time_elapsed.SetFont(wx.Font(
        #     9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))
        # self.txt_time_elapsed.SetForegroundColour(
        #     wx.SystemSettings.GetColour(wx.SYS_COLOUR_3DLIGHT))

        # b_layout_h.Add(self.txt_time_elapsed, 0, wx.ALL, 5)

        main_layout.Add(b_layout_h, 1, wx.EXPAND, 5)

        self.SetSizer(main_layout)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.btn_autoopen.Bind(wx.EVT_CHECKBOX, self.autoOpen)
        self.btn_run.Bind(wx.EVT_BUTTON, self.processData)
        self.btn_openzhpla.Bind(wx.EVT_BUTTON, self.open_zhpla)
        self.btn_openzpdev.Bind(wx.EVT_BUTTON, self.open_zpdev)
        self.btn_openta.Bind(wx.EVT_BUTTON, self.open_talentAnalytics)
        self.btn_openoutputdir.Bind(wx.EVT_BUTTON, self.OpenOutput)

        # Variables
        self.output_zhpla = ''
        self.output_zpdev = ''
        self.output_analytics = ''
        self.output_folder = ''

        self.btn_openzhpla.Disable()
        self.btn_openzpdev.Disable()
        self.btn_openta.Disable()
        self.btn_openoutputdir.Disable()

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def autoOpen(self, event):
        event.Skip()

    def disableUI(self):
        self.zhpla_picker.Disable()
        self.zpdev_picker.Disable()
        self.zpdevretire_picker.Disable()

        self.rdbtn_clean.Disable()
        self.rdbtn_combine.Disable()
        self.rdbtn_cleancombine.Disable()

        self.btn_run.Disable()

        self.btn_autoopen.Disable()

    def enableUI(self):
        self.zhpla_picker.Enable()
        self.zpdev_picker.Enable()
        self.zpdevretire_picker.Enable()

        self.rdbtn_clean.Enable()
        self.rdbtn_combine.Enable()
        self.rdbtn_cleancombine.Enable()

        self.btn_run.Enable()
        self.btn_autoopen.Enable()

    def processData(self, event):
        try:
            self.disableUI()
            self.txt_message.SetLabelText('Running processes...')

            zhplapath = self.zhpla_picker.GetPath()
            zpdevpath = self.zpdev_picker.GetPath()
            zpdev_retirepath = self.zpdevretire_picker.GetPath()

            if self.rdbtn_clean.GetValue():
                operation_type = 1
            elif self.rdbtn_combine.GetValue():
                operation_type = 2
            elif self.rdbtn_cleancombine.GetValue():
                operation_type = 3

            data_dict, operation_type = dc.preProcess(
                zhplapath, zpdevpath, zpdev_retirepath, operation_type)

            df_dictionary, logString = dc.mainProcess(
                data_dict, operation_type)

            for key, value in df_dictionary.items():
                outputfile = (key + dc.current_month).upper() + '.xlsx'

                self.output_zhpla = outputfile if key == 'zhpla' else ''
                self.output_zpdev = outputfile if key == 'zpdev' else ''
                self.output_analytics = outputfile if key == 'analytics' else ''

                value.to_excel(outputfile, index=None)
                if self.btn_autoopen.GetValue():
                    os.startfile(outputfile)

            self.enableOutputButton(df_dictionary)
        except Exception as e:
            print(e)
        finally:
            self.enableUI()
            self.txt_message.SetLabelText(logString)

    def enableOutputButton(self, df_dictionary):
        self.btn_openoutputdir.Enable()
        if 'zhpla' in df_dictionary:
            self.btn_openzhpla.Enable()
        if 'zpdev' in df_dictionary:
            self.btn_openzpdev.Enable()
        if 'analytics' in df_dictionary:
            self.btn_openta.Enable()

    def open_zhpla(self, event):
        os.startfile(self.output_zhpla)

    def open_zpdev(self, event):
        os.startfile(self.output_zpdev)

    def open_talentAnalytics(self, event):
        os.startfile(self.output_analytics)

    def OpenOutput(self, event):
        os.startfile(os.getcwd())


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = AnalyticsFrame(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
