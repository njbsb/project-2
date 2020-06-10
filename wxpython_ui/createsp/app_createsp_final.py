# -*- coding: utf-8 -*-

###########################################################################
# Python code generated with wxFormBuilder (version Oct 26 2018)
# http://www.wxformbuilder.org/
##
# PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import createsp_lib_v5 as cs
import os
import sys
maindirectory = os.getcwd()
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

        self.txt_filelabel = wx.StaticText(input_sz.GetStaticBox(
        ), wx.ID_ANY, u"Select file", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_filelabel.Wrap(-1)

        self.txt_filelabel.SetFont(wx.Font(
            9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))

        input_sz.Add(self.txt_filelabel, 0, wx.ALL, 5)

        self.m_filePicker1 = wx.FilePickerCtrl(input_sz.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        input_sz.Add(self.m_filePicker1, 0, wx.ALL, 5)

        self.txt_mediapath = wx.StaticText(input_sz.GetStaticBox(
        ), wx.ID_ANY, u"Select media path", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_mediapath.Wrap(-1)

        self.txt_mediapath.SetFont(wx.Font(
            9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))

        input_sz.Add(self.txt_mediapath, 0, wx.ALL, 5)

        self.m_dirPicker1 = wx.DirPickerCtrl(input_sz.GetStaticBox(
        ), wx.ID_ANY, wx.EmptyString, u"Select a folder", wx.DefaultPosition, wx.DefaultSize, wx.DIRP_DEFAULT_STYLE)
        self.m_dirPicker1.SetFont(wx.Font(
            9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))

        input_sz.Add(self.m_dirPicker1, 0, wx.ALL, 5)

        content_sz.Add(input_sz, 1, wx.EXPAND, 5)

        output_sz = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, wx.EmptyString), wx.VERTICAL)

        process_sz = wx.StaticBoxSizer(wx.StaticBox(
            output_sz.GetStaticBox(), wx.ID_ANY, u"Process"), wx.VERTICAL)

        self.chbox_autoopen = wx.CheckBox(process_sz.GetStaticBox(
        ), wx.ID_ANY, u"Open when done", wx.DefaultPosition, wx.DefaultSize, 0)
        self.chbox_autoopen.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        self.chbox_autoopen.SetBackgroundColour(wx.Colour(0, 177, 169))

        process_sz.Add(self.chbox_autoopen, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.btn_generate = wx.Button(process_sz.GetStaticBox(
        ), wx.ID_ANY, u"Generate SP Pack", wx.DefaultPosition, wx.DefaultSize, 0)
        process_sz.Add(self.btn_generate, 0, wx.ALIGN_CENTER |
                       wx.BOTTOM | wx.LEFT | wx.RIGHT, 5)

        output_sz.Add(process_sz, 1, wx.ALIGN_CENTER |
                      wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

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
            self, wx.ID_ANY, u"...", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_message.Wrap(-1)

        self.txt_message.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_3DLIGHT))

        mainsizer.Add(self.txt_message, 0, wx.ALL, 5)

        self.SetSizer(mainsizer)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.chbox_autoopen.Bind(wx.EVT_CHECKBOX, self.autoOpen)
        self.btn_generate.Bind(wx.EVT_BUTTON, self.runSP)
        self.btn_openfolder.Bind(wx.EVT_BUTTON, self.openOutputFolder)

        self.btn_openfolder.Disable()

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def autoOpen(self, event):
        event.Skip()

    def runSP(self, event):
        if not self.m_filePicker1.GetPath() == '':
            self.btn_generate.Disable()
            self.m_filePicker1.Disable()
            self.m_dirPicker1.Disable()
            self.chbox_autoopen.Disable()
            filepath = self.m_filePicker1.GetPath()
            mediafolderpath = self.m_dirPicker1.GetPath()
            if mediafolderpath == '':
                self.txt_message.SetLabelText(
                    'Media folder will be automatically be created.')
                import errno
                try:
                    mediafolderpath = os.path.join(maindirectory, 'media')
                    os.makedirs(mediafolderpath)
                except OSError as e:
                    if e.errno != errno.EEXIST:
                        raise
                print(mediafolderpath)

            try:
                # self.txt_message.SetLabelText('Getting data from file...')
                try:
                    filename, sheetname_list = cs.getFileInfo(filepath)
                    data_df_list, id_df_list = cs.get_rawDataframe(
                        filepath, sheetname_list)
                    try:
                        self.txt_message.SetLabelText(
                            'Writing into pptx file...')
                        prs = cs.Presentation()
                        cs.createPPT_data(data_df_list, id_df_list, prs,
                                          sheetname_list, mediafolderpath)
                        try:
                            self.txt_message.SetLabelText('Saving the pack...')
                            sp_filename = filename + '.pptx'
                            outputfile = os.path.join(
                                maindirectory, sp_filename)
                            prs.save(outputfile)
                            if self.chbox_autoopen.IsChecked():
                                os.startfile(outputfile)
                                self.btn_openfolder.Enable()
                                self.txt_message.SetLabelText(
                                    'SP Pack successfully generated.')
                        except Exception as e:
                            self.txt_message.SetLabelText(
                                'Saving failed. Please check.')
                    except Exception as e:
                        self.txt_message.SetLabelText(
                            'Error while writing the data. Please check the data format')
                except Exception as e:
                    self.txt_message.SetLabelText(
                        'Error while reading the file. Please check the data format.')
            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(
                    exc_tb.tb_frame.f_code.co_filename)[1]
                print('Error')
                print(exc_type, fname, exc_tb.tb_lineno)
                self.txt_message.SetLabelText(str(e))
            finally:
                self.btn_generate.Enable()
                self.m_filePicker1.Enable()
                self.m_dirPicker1.Enable()
                self.chbox_autoopen.Enable()
                self.chbox_autoopen.SetValue(False)
        else:
            self.txt_message.SetLabelText(
                'No file selected. Please select one SP Excel file.')

    def openOutputFolder(self, event):
        os.startfile(maindirectory)


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = mainframe(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
