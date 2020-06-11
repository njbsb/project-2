# -*- coding: utf-8 -*-

###########################################################################
# Python code generated with wxFormBuilder (version Oct 26 2018)
# http://www.wxformbuilder.org/
##
# PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

from datetime import datetime
import wx
import wx.xrc
import pandas as pd
import os
import cleanzh_lib as cleanzh
import cleanzp_lib as cleanzp
import combinedb_lib as combinedb
import time
from datetime import datetime
###########################################################################
# Class frame_ta
###########################################################################
wildcard = "Excel File (*.xlsx)|*.xlsx|" \
    "All files (*.*)|*.*"
currentMonth = str(datetime.now().strftime('%B')).upper()
# delaytime = 1
zh_columns = ['Pos ID', 'End Date', 'Position', 'Vacancy Status', 'Email Address', 'Mirror Pos ID', 'Mirror Position', 'Superior Pos ID', 'Sec ID', 'Section', 'Dept ID', 'Department', 'Division ID', 'Division', 'Sector ID', 'Sector', 'Comp. ID Position', 'Comp. Position', 'Buss ID', 'Business', 'C.Center', 'EG Post. ID', 'EG Position', 'ESG Post. ID', 'ESG Position', 'PT1', 'PT2', 'Staff No', 'Staff Name',
              'Overseas Staff No', 'Overseas Staff Name', 'Overseas Staff Comp. ID', 'Overseas Staff Comp.', 'Staff Comp. ID', 'Staff Comp.', 'Staff EG ID', 'Staff EG', 'Staff ESG ID', 'Staff ESG', 'Staff Work Contract ID', 'Staff Work Contract', 'Gender', 'Race', 'SG', 'Lvl', 'Tier', 'JG', 'Est.JG', 'EQV JG', 'OLIVE JG', 'Zone', 'Home SKG', 'Pos.SKG', 'JobID(C)', 'Level', 'NE Area', 'Loc ID', 'Location Text', 'Vac Date', 'Start Date']
zp_columns = ['Personnel Number', 'Formatted Name of Employee or Applicant', 'Gender Key', 'Gender Key Desc', 'Date of Birth', 'Country of Birth', 'Country of Birth Desc', 'State', 'State Desc', 'Birthplace', 'Nationality', 'Nationality Desc', 'Join Date', 'Join Dept', 'Sal. Grade', 'Lvl', 'Job Grade',
              'Position', 'Company', 'Sector', 'Division', 'Department', 'Section', 'SG Since', 'Pos. SKG', 'Work Center', 'Home SKG', 'Date Post', 'Join Div.', 'Join Comp.', 'PPA-19', 'Cluster', 'PPA-18', 'Cluster', 'PPA-17', 'Cluster', 'PPA-16', 'Cluster', 'PPA-15', 'Cluster', 'Task Type', 'Task Type Desc', 'Date of Task']


def clean_method(dbtype, inputpath):
    print('Running cleanup on', dbtype)
    z_df = pd.read_excel(inputpath, skiprows=4, nrows=None)
    if dbtype == 'zhpla':
        cleandf = cleanzh.clean_rawdf(z_df)
        z_columns = zh_columns
    else:  # for zpdev
        cleandf = cleanzp.clean_rawdf(z_df)
        z_columns = zp_columns
    list_colname = [colname[0] for i, colname in cleandf.iteritems()]
    checkColumnExist = all(item in list_colname for item in z_columns)
    return cleandf, checkColumnExist, list_colname


def time_elapsed(seconds):
    totaltime = str(time.strftime("%H:%M:%S", time.gmtime(seconds)))
    print(totaltime)
    timeHR = time.strftime("%H", time.gmtime(seconds)).lstrip('0')
    timeMT = time.strftime("%M", time.gmtime(seconds)).lstrip('0')
    timeSC = time.strftime("%S", time.gmtime(seconds)).lstrip('0')
    # print(timeHR, timeMT, timeSC)  # type str
    timevalue = [timeHR, timeMT, timeSC]
    timename = ['hour', 'minute', 'second']
    str_time = ''
    for tv, tn in zip(timevalue, timename):
        if not tv == '':
            if int(tv) > 1:
                tn += 's'
            str_time += tv + ' ' + tn + ' '
    print(str_time)
    return str_time


class frame_ta (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"TTM: Talent Analytics", pos=wx.DefaultPosition, size=wx.Size(
            600, 480), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        main_layout = wx.BoxSizer(wx.VERTICAL)

        self.txt_title = wx.StaticText(self, wx.ID_ANY, u"Clean ZHPLA/ZPDEV\nTalent Analytics",
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

        self.checkbox_zhpla = wx.CheckBox(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"ZHPLA", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkbox_zhpla.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        self.checkbox_zhpla.SetBackgroundColour(wx.Colour(0, 177, 169))

        layout_input.Add(self.checkbox_zhpla, 0, wx.ALL, 5)

        self.btn_selectzhpla = wx.Button(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"ZHPLA...", wx.DefaultPosition, wx.Size(-1, -1), 0)
        layout_input.Add(self.btn_selectzhpla, 0, wx.ALL, 5)

        self.txt_zhplafile = wx.StaticText(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"zhpla_file.xlsx", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_zhplafile.Wrap(-1)

        layout_input.Add(self.txt_zhplafile, 0, wx.ALL, 5)

        self.checkbox_zpdev = wx.CheckBox(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"ZPDEV", wx.DefaultPosition, wx.DefaultSize, 0)
        self.checkbox_zpdev.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        self.checkbox_zpdev.SetBackgroundColour(wx.Colour(0, 177, 169))

        layout_input.Add(self.checkbox_zpdev, 0, wx.ALL, 5)
        # self.btn_selectzhpla = wx.Button(layout_input.GetStaticBox(
        # ), wx.ID_ANY, u"ZHPLA...", wx.DefaultPosition, wx.Size(-1, -1), 0)
        # layout_input.Add(self.btn_selectzhpla, 0, wx.ALL, 5)

        # self.txt_zhplafile = wx.StaticText(layout_input.GetStaticBox(
        # ), wx.ID_ANY, u"zhplafile.xls", wx.DefaultPosition, wx.DefaultSize, 0)
        # self.txt_zhplafile.Wrap(-1)

        # layout_input.Add(self.txt_zhplafile, 0, wx.ALL, 5)

        self.btn_selectzpdev = wx.Button(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"ZPDEV...", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_input.Add(self.btn_selectzpdev, 0, wx.ALL, 5)

        self.txt_zpdevfile = wx.StaticText(layout_input.GetStaticBox(
        ), wx.ID_ANY, u"zpdev_file.xlsx", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_zpdevfile.Wrap(-1)

        layout_input.Add(self.txt_zpdevfile, 0, wx.ALL, 5)

        b_layout_p1.Add(layout_input, 1, wx.EXPAND, 5)

        layout_process = wx.StaticBoxSizer(wx.StaticBox(
            b_layout_p1.GetStaticBox(), wx.ID_ANY, u"select process"), wx.VERTICAL)

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

        self.chbox_autoopen = wx.CheckBox(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Open when done", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_choice.Add(self.chbox_autoopen, 0, wx.ALL, 5)

        layout_process.Add(layout_choice, 1, wx.EXPAND, 5)

        layout_button = wx.BoxSizer(wx.VERTICAL)

        # self.btn_run = wx.Button(layout_process.GetStaticBox(
        # ), wx.ID_ANY, u"text", wx.DefaultPosition, wx.DefaultSize, 0)
        self.btn_run = wx.Button(layout_process.GetStaticBox(
        ), wx.ID_ANY, u"Run", wx.DefaultPosition, wx.Size(120, -1), 0)
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
            b_layout_p1.GetStaticBox(), wx.ID_ANY, u"output"), wx.HORIZONTAL)

        layout_openfile = wx.BoxSizer(wx.VERTICAL)

        self.btn_openzhpla = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"zhpla", wx.DefaultPosition, wx.Size(60, -1), 0)
        layout_openfile.Add(self.btn_openzhpla, 0, wx.ALL, 5)

        self.btn_openzpdev = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"zpdev", wx.DefaultPosition, wx.Size(60, -1), 0)
        layout_openfile.Add(self.btn_openzpdev, 0, wx.ALL, 5)

        self.btn_openta = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"talent analytics", wx.DefaultPosition, wx.Size(-1, -1), 0)
        layout_openfile.Add(self.btn_openta, 0, wx.ALL, 5)

        layout_output.Add(layout_openfile, 1, wx.ALIGN_CENTER | wx.EXPAND, 5)

        self.btn_openoutputdir = wx.Button(layout_output.GetStaticBox(
        ), wx.ID_ANY, u"Open folder", wx.DefaultPosition, wx.DefaultSize, 0)
        layout_output.Add(self.btn_openoutputdir, 0, wx.ALL, 5)

        b_layout_p1.Add(layout_output, 1, wx.EXPAND, 5)

        b_layout_h.Add(b_layout_p1, 1, wx.EXPAND, 5)

        self.txt_message = wx.StaticText(
            self, wx.ID_ANY, u"", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_message.Wrap(-1)

        self.txt_message.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        b_layout_h.Add(self.txt_message, 0, wx.ALL, 5)

        self.txt_time_elapsed = wx.StaticText(
            self, wx.ID_ANY, u"Time elapsed:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.txt_time_elapsed.Wrap(-1)

        self.txt_time_elapsed.SetFont(wx.Font(
            9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))
        self.txt_time_elapsed.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_3DLIGHT))

        b_layout_h.Add(self.txt_time_elapsed, 0, wx.ALL, 5)

        main_layout.Add(b_layout_h, 1, wx.EXPAND, 5)

        self.SetSizer(main_layout)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.checkbox_zhpla.Bind(wx.EVT_CHECKBOX, self.tickZhpla)
        self.btn_selectzhpla.Bind(wx.EVT_BUTTON, self.selectZhpla)
        self.checkbox_zpdev.Bind(wx.EVT_CHECKBOX, self.tickZpdev)
        self.btn_selectzpdev.Bind(wx.EVT_BUTTON, self.selectZpdev)
        self.rdbtn_clean.Bind(wx.EVT_RADIOBUTTON, self.showClean)
        self.rdbtn_combine.Bind(wx.EVT_RADIOBUTTON, self.showCombine)
        self.rdbtn_cleancombine.Bind(wx.EVT_RADIOBUTTON, self.showCleanCombine)
        self.chbox_autoopen.Bind(wx.EVT_CHECKBOX, self.autoOpen)
        self.btn_run.Bind(wx.EVT_BUTTON, self.action)
        self.btn_openzhpla.Bind(wx.EVT_BUTTON, self.open_zhpla)
        self.btn_openzpdev.Bind(wx.EVT_BUTTON, self.open_zpdev)
        self.btn_openta.Bind(wx.EVT_BUTTON, self.open_talentAnalytics)
        self.btn_openoutputdir.Bind(wx.EVT_BUTTON, self.OpenOutput)

        # Enable Disable buttons
        self.btn_openoutputdir.Disable()
        self.btn_openta.Disable()
        self.btn_openzhpla.Disable()
        self.btn_openzpdev.Disable()
        self.btn_run.Disable()
        self.btn_selectzhpla.Disable()
        self.btn_selectzpdev.Disable()
        self.rdbtn_cleancombine.Disable()
        # self.rdbtn_cleancombine.Disable()
        # self.rdbtn_combine.Disable()
        self.txt_time_elapsed.SetLabelText(
            'Current time: '+str(datetime.now().strftime("%H:%M")))

        # Common variable
        self.currentDirectory = os.path.dirname(__file__)
        self.zhpla_inputpath = ''
        self.zpdev_inputpath = ''
        zhpla_outfile = 'ZHPLA_' + currentMonth + '.xlsx'
        zpdev_outfile = 'ZPDEV_' + currentMonth + '.xlsx'
        analyt_outfile = 'TALENT_ANALYTICS_' + currentMonth + '.xlsx'

        self.zhpla_outputpath = os.path.join(
            self.currentDirectory, zhpla_outfile)
        self.zpdev_outputpath = os.path.join(
            self.currentDirectory, zpdev_outfile)
        self.analytics_outputpath = os.path.join(
            self.currentDirectory, analyt_outfile)

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def tickZhpla(self, event):
        if(self.btn_selectzhpla.IsEnabled()):
            self.btn_selectzhpla.Disable()
            self.txt_zhplafile.SetLabelText('zhplafile.xlsx')
            self.zhpla_inputpath = ''
            self.txt_message.SetLabelText('')
        else:
            self.btn_selectzhpla.Enable()

    def selectZhpla(self, event):
        dlg = wx.FileDialog(self, message="Choose a file (ZHPLA)", defaultDir=self.currentDirectory,
                            defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            filename = dlg.GetFilename()
            print("You chose the following file(s): %s" % filename)
            self.txt_zhplafile.SetLabelText(filename)
            pa = ''.join(paths)
            self.zhpla_inputpath = pa
        dlg.Destroy()

    def tickZpdev(self, event):
        if(self.btn_selectzpdev.IsEnabled()):
            self.btn_selectzpdev.Disable()
            self.txt_zpdevfile.SetLabelText('zpdevfile.xlsx')
            self.zpdev_inputpath = ''
            self.txt_message.SetLabelText('')
        else:
            self.btn_selectzpdev.Enable()

    def selectZpdev(self, event):
        dlg = wx.FileDialog(self, message="Choose a file (ZPDEV)", defaultDir=self.currentDirectory,
                            defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            filename = dlg.GetFilename()
            print("You chose the following file(s): %s" % filename)
            self.txt_zpdevfile.SetLabelText(filename)
            pa = ''.join(paths)
            self.zpdev_inputpath = pa
        dlg.Destroy()

    def showClean(self, event):
        self.btn_run.SetLabelText('Clean')
        self.btn_run.Enable()

    def showCombine(self, event):
        self.btn_run.SetLabelText('Combine')
        self.btn_run.Enable()

    def showCleanCombine(self, event):
        self.btn_run.SetLabelText('Clean & Combine')
        self.btn_run.Enable()

    def action(self, event):
        actiontypetext = self.btn_run.GetLabelText()

        if not self.zhpla_inputpath == '' or not self.zpdev_inputpath == '':
            self.btn_selectzhpla.Disable()
            self.btn_selectzpdev.Disable()
            starttime = time.time()
            # self.txt_status.SetLabelText('Running')
            if actiontypetext == 'Clean':
                if not self.zhpla_inputpath == '' or not self.zpdev_inputpath == '':
                    try:
                        self.btn_run.Disable()
                        now = datetime.now()
                        current_time = now.strftime("%H:%M:%S")
                        # print("Current Time =", current_time)
                        self.txt_time_elapsed.SetLabelText(current_time)
                        if not self.zhpla_inputpath == '':
                            try:
                                self.txt_message.SetLabelText(
                                    'Running cleanup on zhpla...')
                                cleandf, checkColumnExist, list_colname = clean_method(
                                    'zhpla', self. zhpla_inputpath)
                                if checkColumnExist:  # if true
                                    finaldf = cleanzh.format_zhpla(cleandf)
                                    finaldf.to_excel(self.zhpla_outputpath,
                                                     index=None)
                                    print('ZH done')
                                    self.txt_message.SetLabelText(
                                        'ZHPLA cleaned up successfully')
                                    if(self.chbox_autoopen.GetValue()):
                                        os.startfile(self.zhpla_outputpath)
                                else:
                                    self.txt_message.SetLabelText(
                                        'Some columns in zhpla file is missing: ' + str(set(zh_columns).difference(list_colname)))
                                # self.checkbox_zhpla.SetValue(False)
                            except Exception as e:
                                self.txt_message.SetLabelText(
                                    'Zhpla error: ' + str(e))
                                print(e)
                            else:
                                self.btn_openzhpla.Enable()
                        if not self.zpdev_inputpath == '':
                            try:
                                # ZPDEV
                                self.txt_message.SetLabelText(
                                    'Running cleanup on zpdev...')
                                cleandf, checkColumnExist, list_colname = clean_method(
                                    'zpdev', self.zpdev_inputpath)
                                if checkColumnExist:
                                    finaldf = cleanzp.format_zpdev(cleandf)
                                    finaldf.to_excel(self.zpdev_outputpath,
                                                     index=None)
                                    self.txt_message.SetLabelText(
                                        'ZPDEV cleaned up successfully')
                                    if(self.chbox_autoopen.GetValue()):
                                        os.startfile(self.zpdev_outputpath)
                                    print('ZP done')
                                else:
                                    self.txt_message.SetLabelText(
                                        'Some columns in zpdev file is missing: ' + str(set(zp_columns).difference(list_colname)))
                                # self.checkbox_zpdev.SetValue(False)
                            except Exception as e:
                                self.txt_message.SetLabelText(
                                    'Zpdev error: ' + str(e))
                                print(e)
                            else:
                                self.btn_openzpdev.Enable()
                    except Exception as e:
                        self.txt_message.SetLabelText(str(e))
                    else:
                        self.btn_openoutputdir.Enable()
                    finally:
                        self.checkbox_zhpla.SetValue(False)
                        self.checkbox_zpdev.SetValue(False)
                        self.txt_zhplafile.SetLabelText('')
                        self.txt_zpdevfile.SetLabelText('')
                        self.zhpla_inputpath = ''
                        self.zpdev_inputpath = ''
                        self.rdbtn_clean.SetValue(False)
                        self.rdbtn_combine.SetValue(False)
                        self.chbox_autoopen.SetValue(False)
                        self.btn_run.Disable()
                        self.btn_run.SetLabelText('Run')
                else:
                    self.txt_message.SetLabelText(
                        'Please select at least one database file first!')

            elif actiontypetext == 'Combine':
                if not self.zhpla_inputpath == '' and not self.zpdev_inputpath == '':
                    try:
                        print('Combining the databases')
                        self.txt_message.SetLabelText('Running...')
                        self.btn_run.Disable()
                        # self.checkbox_zhpla.Disable()
                        # self.checkbox_zpdev.Disable()
                        zh_df = pd.read_excel(self.zhpla_inputpath)
                        zp_df = pd.read_excel(self.zpdev_inputpath)
                        zh_df, zp_df = combinedb.prepareDataframe(zh_df, zp_df)
                        ta_df = combinedb.merge_database(zh_df, zp_df)
                        ta_df.to_excel(self.analytics_outputpath, index=False)
                        self.txt_message.SetLabelText(
                            'Database combine has successfully executed')
                        self.btn_openta.Enable()
                        if(self.chbox_autoopen.GetValue()):
                            os.startfile(self.analytics_outputpath)
                    except Exception as e:
                        self.txt_message.SetLabelText('Error' + str(e))
                        print('Error', e)
                        self.checkbox_zhpla.SetValue(False)
                        self.checkbox_zpdev.SetValue(False)
                        self.rdbtn_combine.SetValue(False)
                    else:
                        self.btn_openoutputdir.Enable()
                        self.btn_openta.Enable()
                    finally:
                        self.checkbox_zhpla.Enable()
                        self.checkbox_zpdev.Enable()
                        self.checkbox_zhpla.SetValue(False)
                        self.checkbox_zpdev.SetValue(False)
                        self.txt_zhplafile.SetLabelText('')
                        self.txt_zpdevfile.SetLabelText('')
                        self.rdbtn_clean.SetValue(False)
                        self.rdbtn_combine.SetValue(False)
                        self.chbox_autoopen.SetValue(False)
                        self.btn_run.Disable()
                        self.btn_run.SetLabelText('Run')
                else:
                    self.txt_message.SetLabelText(
                        'You need to select both databases to run combine function!')
            else:
                if not self.zhpla_inputpath == '' and not self.zpdev_inputpath == '':
                    pass
                else:
                    self.txt_message.SetLabelText(
                        'You need to select both databases to run combine function!')
            endtime = time.time()
            totalseconds = int(endtime - starttime)
            str_time = time_elapsed(totalseconds)
            self.txt_time_elapsed.SetLabelText(
                'Time elapsed: ' + str_time)
        else:
            self.txt_message.SetLabelText(
                'Please select at least one database file first!')

    def autoOpen(self, event):
        event.Skip()

    def open_zhpla(self, event):
        os.startfile(self.zhpla_outputpath)

    def open_zpdev(self, event):
        os.startfile(self.zpdev_outputpath)

    def open_talentAnalytics(self, event):
        os.startfile(self.analytics_outputpath)

    def OpenOutput(self, event):
        os.startfile(self.currentDirectory)


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = frame_ta(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
