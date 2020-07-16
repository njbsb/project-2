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
import july_sp as sp
###########################################################################
# Class spFrame
###########################################################################


class spFrame (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"TTM: Succession Planner", pos=wx.DefaultPosition, size=wx.Size(
            500, 300), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        mainsizer = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(
            self, wx.ID_ANY, u"Succession Planner", wx.DefaultPosition, wx.DefaultSize, 0)
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

        self.txt_filelabel = wx.StaticText(input_sz.GetStaticBox(
        ), wx.ID_ANY, u"Select file", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_filelabel.Wrap(-1)

        self.txt_filelabel.SetFont(wx.Font(
            9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))

        input_sz.Add(self.txt_filelabel, 0, wx.ALL, 5)

        self.sp_picker = wx.FilePickerCtrl(input_sz.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        input_sz.Add(self.sp_picker, 0, wx.ALL, 5)

        self.txt_mediapath = wx.StaticText(input_sz.GetStaticBox(
        ), wx.ID_ANY, u"Select media path", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_mediapath.Wrap(-1)

        self.txt_mediapath.SetFont(wx.Font(
            9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))

        input_sz.Add(self.txt_mediapath, 0, wx.ALL, 5)

        self.mediafolder_picker = wx.DirPickerCtrl(input_sz.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, u"Select a folder", wx.DefaultPosition, wx.DefaultSize, wx.DIRP_DEFAULT_STYLE)
        self.mediafolder_picker.SetFont(wx.Font(
            9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))

        input_sz.Add(self.mediafolder_picker, 0, wx.ALL, 5)

        content_sz.Add(input_sz, 1, wx.EXPAND, 5)

        output_sz = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, wx.EmptyString), wx.VERTICAL)

        process_sz = wx.StaticBoxSizer(wx.StaticBox(
            output_sz.GetStaticBox(), wx.ID_ANY, u"Process"), wx.VERTICAL)

        self.chbox_autoopen = wx.CheckBox(process_sz.GetStaticBox(
        ), wx.ID_ANY, u"Open when done", wx.DefaultPosition, wx.DefaultSize, 0)
        self.chbox_autoopen.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWTEXT))
        self.chbox_autoopen.SetBackgroundColour(wx.Colour(0, 177, 169))

        process_sz.Add(self.chbox_autoopen, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.btn_generate = wx.Button(process_sz.GetStaticBox(
        ), wx.ID_ANY, u"Generate SP Pack", wx.DefaultPosition, wx.DefaultSize, 0)
        process_sz.Add(self.btn_generate, 0, wx.ALIGN_CENTER |
                       wx.BOTTOM | wx.LEFT | wx.RIGHT, 5)

        output_sz.Add(process_sz, 1, wx.EXPAND, 5)

        output_sizer = wx.StaticBoxSizer(wx.StaticBox(
            output_sz.GetStaticBox(), wx.ID_ANY, u"Output"), wx.VERTICAL)

        self.btn_openfolder = wx.Button(output_sizer.GetStaticBox(
        ), wx.ID_ANY, u"Open folder", wx.DefaultPosition, wx.DefaultSize, 0)
        output_sizer.Add(self.btn_openfolder, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        output_sz.Add(output_sizer, 1, wx.EXPAND, 5)

        content_sz.Add(output_sz, 1, wx.EXPAND, 5)

        mainsizer.Add(content_sz, 1, wx.EXPAND |
                      wx.LEFT | wx.RIGHT | wx.TOP, 5)

        self.txt_message = wx.StaticText(
            self, wx.ID_ANY, u"Welcome to Succession Planner", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_message.Wrap(-1)

        self.txt_message.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_3DLIGHT))

        mainsizer.Add(self.txt_message, 0, wx.ALL, 5)

        self.SetSizer(mainsizer)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.btn_generate.Bind(wx.EVT_BUTTON, self.generate_sp)
        self.btn_openfolder.Bind(wx.EVT_BUTTON, self.openOutputFolder)

        # Variable
        self.outputfolder = os.getcwd()

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def generate_sp(self, event):
        sp_path = self.sp_picker.GetPath()
        mediafolder = self.mediafolder_picker.GetPath()
        if sp_path == '':
            self.txt_message.SetLabelText('Please select an sp file first')
        else:
            presentation, outputfile = sp.mainProcess(sp_path, mediafolder)
            if self.chbox_autoopen.GetValue():
                os.startfile(outputfile)

    def openOutputFolder(self, event):
        os.startfile(self.outputfolder)


###########################################################################
# Class otherFrame
###########################################################################

class otherFrame (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=wx.EmptyString, pos=wx.DefaultPosition, size=wx.Size(
            500, 300), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        mSizer = wx.BoxSizer(wx.VERTICAL)

        self.SetSizer(mSizer)
        self.Layout()

        self.Centre(wx.BOTH)

    def __del__(self):
        pass


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = spFrame(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
