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
import re
import PyPDF2
maindirectory = os.path.dirname(__file__)
###########################################################################
# Class main_frame
###########################################################################


def getStaffID(report_type, page):
    print(report_type)
    if(report_type == 'BePCB' or report_type == 'LC'):
        assID = 'Assessee ID:'
        nextWord = 'Position'
    elif(report_type == 'CV'):
        assID = 'STAFF NO.: '
        nextWord = 'POSITION'
    # assID = ''
    assID_len = len(assID)
    text = page.extractText()
    asid_id = text.find(assID)
    nextword_index = text.find(nextWord)
    start = asid_id + assID_len
    end = nextword_index
    staffID = text[start:end].lstrip("0")
    return staffID


def getPageList(pdf, keyword):
    pagecount = pdf.getNumPages()
    print("Number of pages: {}".format(pagecount))
    # String = "STAFF DETAILS"
    pagelist = []
    for i in range(0, pagecount):
        page = pdf.getPage(i)
        text = page.extractText()
        ResSearch = re.search(keyword, text)  # String represents keyword
        if(ResSearch != None):
            # insert page num of the first page of each report into list
            pagelist.append(i)
    return pagelist, pagecount


def splitPdf(pdf, report_type, pagelist, pagecount, outputfolder):
    for i, startpage in enumerate(pagelist):
        output = PyPDF2.PdfFileWriter()
        if(i < len(pagelist) - 1):
            pagelength = pagelist[i+1] - startpage
        else:
            pagelength = pagecount - startpage
        # pagelength is the number of page for each report
        for page in range(pagelength):
            pagenum = startpage + page
            output.addPage(pdf.getPage(pagenum))
        # for name
        firstPage = pdf.getPage(startpage)
        # put if statement if bepcb, then getstaffDID
        staffID = getStaffID(report_type, firstPage)
        filename = "%s.pdf" % staffID
        # output dir : output folder
        outputdir = os.path.join(maindirectory, outputfolder)
        if not os.path.exists(outputdir):
            os.makedirs(outputdir)
        outputDir = outputdir
        os.startfile(outputdir)
        with open(outputdir + "/" + filename, "wb") as outputStream:
            output.write(outputStream)
    print("DONE")


class main_frame (wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"SplitPDF - Report", pos=wx.DefaultPosition,
                          size=wx.Size(500, 326), style=wx.CAPTION | wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(wx.NORMAL_FONT.GetPointSize(), wx.FONTFAMILY_DEFAULT,
                             wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Century Gothic"))
        self.SetBackgroundColour(wx.Colour(0, 177, 169))

        boxSizerMain = wx.BoxSizer(wx.VERTICAL)

        self.text_title = wx.StaticText(
            self, wx.ID_ANY, u"SplitPDF - Report", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_title.Wrap(-1)

        self.text_title.SetFont(wx.Font(
            14, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "Century Gothic"))
        self.text_title.SetForegroundColour(
            wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        boxSizerMain.Add(self.text_title, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        bsizer_input = wx.BoxSizer(wx.HORIZONTAL)

        sbsizer_src = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Source file"), wx.VERTICAL)

        bsizer_src = wx.BoxSizer(wx.VERTICAL)

        self.filepicker = wx.FilePickerCtrl(sbsizer_src.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a PDF file",
                                            u"PDF source (*.pdf)|*.pdf|"
                                            "All files (*.*)|*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
        bsizer_src.Add(self.filepicker, 0, wx.ALL, 6)

        self.label_filename = wx.StaticText(sbsizer_src.GetStaticBox(
        ), wx.ID_ANY, u"file_name.pdf ?", wx.DefaultPosition, wx.DefaultSize, 0)
        self.label_filename.Wrap(-1)

        bsizer_src.Add(self.label_filename, 0, wx.ALL, 5)

        sbsizer_src.Add(bsizer_src, 1, wx.EXPAND, 5)

        bsizer_input.Add(sbsizer_src, 1, wx.EXPAND, 5)

        sbsizer_txtfld = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Report Type"), wx.VERTICAL)

        cb_reporttypeChoices = [u"BePCB", u"CV", u"LC"]
        self.cb_reporttype = wx.ComboBox(sbsizer_txtfld.GetStaticBox(
        ), wx.ID_ANY, u"Select Type", wx.DefaultPosition, wx.DefaultSize, cb_reporttypeChoices, 0)
        sbsizer_txtfld.Add(self.cb_reporttype, 0, wx.ALL | wx.EXPAND, 5)

        bsizer_input.Add(sbsizer_txtfld, 1, wx.EXPAND, 5)

        boxSizerMain.Add(bsizer_input, 1, wx.ALL | wx.EXPAND, 5)

        self.btn_split = wx.Button(
            self, wx.ID_ANY, u"Split!", wx.DefaultPosition, wx.DefaultSize, 0)
        boxSizerMain.Add(self.btn_split, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        sbsizer_output = wx.StaticBoxSizer(wx.StaticBox(
            self, wx.ID_ANY, u"Output"), wx.HORIZONTAL)

        bsizer_txt = wx.BoxSizer(wx.VERTICAL)

        self.text_status = wx.StaticText(sbsizer_output.GetStaticBox(
        ), wx.ID_ANY, u"Status:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_status.Wrap(-1)

        bsizer_txt.Add(self.text_status, 0, wx.ALL, 5)

        self.text_reportcount = wx.StaticText(sbsizer_output.GetStaticBox(
        ), wx.ID_ANY, u"Report Count:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_reportcount.Wrap(-1)

        bsizer_txt.Add(self.text_reportcount, 0, wx.ALL, 5)

        sbsizer_output.Add(bsizer_txt, 1, wx.EXPAND, 5)

        bsizer_btnopen = wx.BoxSizer(wx.VERTICAL)

        self.btn_opendir = wx.Button(sbsizer_output.GetStaticBox(
        ), wx.ID_ANY, u"Output Directory...", wx.DefaultPosition, wx.DefaultSize, 0)
        bsizer_btnopen.Add(self.btn_opendir, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        sbsizer_output.Add(bsizer_btnopen, 1, wx.EXPAND, 5)

        boxSizerMain.Add(sbsizer_output, 1, wx.ALL | wx.EXPAND, 5)

        self.SetSizer(boxSizerMain)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.filepicker.Bind(wx.EVT_FILEPICKER_CHANGED, self.changeFileLabel)
        self.btn_split.Bind(wx.EVT_BUTTON, self.splitFile)
        self.btn_opendir.Bind(wx.EVT_BUTTON, self.openDirectory)

        # Variables
        self.outputfolder = ''
        self.btn_opendir.Disable()
        self.btn_split.Disable()

    def __del__(self):
        pass

    # Virtual event handlers, overide them in your derived class
    def changeFileLabel(self, event):
        filepath = self.filepicker.GetPath()
        print(filepath)
        filename = os.path.basename(filepath)
        self.label_filename.SetLabelText(str(filename))
        if not filepath == '':
            self.btn_split.Enable()
        else:
            self.btn_split.Disable()

    def splitFile(self, event):
        keywordlist = ['Assessee ID:', 'STAFF NO.:', 'Assessee ID:']
        input_path = self.filepicker.GetPath()
        report_type = self.cb_reporttype.GetValue()
        if not input_path and report_type == 'Select Type':
            self.text_status.SetLabelText(
                'You have not selected any report and filetype')
        else:
            try:
                self.filepicker.Disable()
                self.btn_split.Disable()
                self.cb_reporttype.Disable()

                self.outputfolder = os.path.join(
                    maindirectory, report_type.lower())
                object = PyPDF2.PdfFileReader(input_path)
                keyword = keywordlist[self.cb_reporttype.GetSelection()]
                try:
                    pagelist, pagecount = getPageList(object, keyword)
                    print(pagelist)
                    try:
                        splitPdf(object, report_type, pagelist,
                                 pagecount, self.outputfolder)
                        self.text_status.SetLabelText('Status: successful')
                        self.text_reportcount.SetLabelText(
                            'Report count: ' + str(len(pagelist)))
                        self.btn_opendir.Enable()
                    except Exception as e:
                        print(str(e))
                        self.text_status.SetLabelText(
                            'Error while splitting the report.')
                except Exception as e:
                    print(str(e))
                    self.text_status.SetLabelText(
                        'Error while reading the report.')
            except Exception as e:
                print(str(e))
                self.text_status.SetLabelText(
                    'Error while reading the report.')
            finally:
                self.filepicker.Enable()
                self.btn_split.Enable()
                self.cb_reporttype.Enable()

    def openDirectory(self, event):
        os.startfile(self.outputfolder)


class MainApp(wx.App):
    def OnInit(self):
        mainFrame = main_frame(None)
        mainFrame.Show(True)
        return True


if __name__ == '__main__':
    app = MainApp()
    app.MainLoop()
