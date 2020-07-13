# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Oct 26 2018)
## http://www.wxformbuilder.org/
##
## PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc

###########################################################################
## Class dashboardFrame
###########################################################################

class dashboardFrame ( wx.Frame ):

	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"TM: TOTAL Dashboard", pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE )

		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
		self.SetFont( wx.Font( 10, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Museo 500" ) )
		self.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )

		mainsizer = wx.BoxSizer( wx.VERTICAL )

		self.title = wx.StaticText( self, wx.ID_ANY, u"TOTAL Dashboard", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
		self.title.Wrap( -1 )

		self.title.SetFont( wx.Font( 14, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, True, "Museo 500" ) )
		self.title.SetForegroundColour( wx.Colour( 255, 255, 255 ) )

		mainsizer.Add( self.title, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		gridSizer = wx.GridSizer( 0, 2, 0, 0 )

		bsizer1 = wx.BoxSizer( wx.VERTICAL )

		self.title_analytics = wx.StaticText( self, wx.ID_ANY, u"Talent Analytics", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
		self.title_analytics.Wrap( -1 )

		bsizer1.Add( self.title_analytics, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		self.btn_analytics = wx.Button( self, wx.ID_ANY, u"Talent Analytics", wx.DefaultPosition, wx.DefaultSize, 0 )
		bsizer1.Add( self.btn_analytics, 0, wx.ALL, 5 )


		gridSizer.Add( bsizer1, 1, wx.ALIGN_CENTER|wx.ALL, 5 )

		bsizer2 = wx.BoxSizer( wx.VERTICAL )

		self.title_sp = wx.StaticText( self, wx.ID_ANY, u"Succession Planner", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.title_sp.Wrap( -1 )

		bsizer2.Add( self.title_sp, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		self.btn_splanner = wx.Button( self, wx.ID_ANY, u"Succession Planner", wx.DefaultPosition, wx.DefaultSize, 0 )
		bsizer2.Add( self.btn_splanner, 0, wx.ALL, 5 )


		gridSizer.Add( bsizer2, 1, wx.ALIGN_CENTER|wx.ALL, 5 )

		bsizer3 = wx.BoxSizer( wx.VERTICAL )

		self.title_reporter = wx.StaticText( self, wx.ID_ANY, u"Reporter", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.title_reporter.Wrap( -1 )

		bsizer3.Add( self.title_reporter, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		self.btn_reporter = wx.Button( self, wx.ID_ANY, u"Reporter", wx.DefaultPosition, wx.DefaultSize, 0 )
		bsizer3.Add( self.btn_reporter, 0, wx.ALL, 5 )


		gridSizer.Add( bsizer3, 1, wx.ALIGN_CENTER|wx.ALL, 5 )

		bsizer4 = wx.BoxSizer( wx.VERTICAL )

		self.title_talentdoc = wx.StaticText( self, wx.ID_ANY, u"TalentDocs", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.title_talentdoc.Wrap( -1 )

		bsizer4.Add( self.title_talentdoc, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		self.btn_talentdoc = wx.Button( self, wx.ID_ANY, u"TalentDocs", wx.DefaultPosition, wx.DefaultSize, 0 )
		bsizer4.Add( self.btn_talentdoc, 0, wx.ALL, 5 )


		gridSizer.Add( bsizer4, 1, wx.ALIGN_CENTER, 5 )


		mainsizer.Add( gridSizer, 1, wx.EXPAND, 5 )

		self.m_staticText7 = wx.StaticText( self, wx.ID_ANY, u"for Top Talent Management, PETRONAS", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
		self.m_staticText7.Wrap( -1 )

		self.m_staticText7.SetFont( wx.Font( 8, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Museo 500" ) )

		mainsizer.Add( self.m_staticText7, 0, wx.ALIGN_CENTER|wx.ALL, 5 )


		self.SetSizer( mainsizer )
		self.Layout()

		self.Centre( wx.BOTH )

		# Connect Events
		self.btn_analytics.Bind( wx.EVT_BUTTON, self.openframe_analytics )
		self.btn_splanner.Bind( wx.EVT_BUTTON, self.openframe_splanner )
		self.btn_reporter.Bind( wx.EVT_BUTTON, self.openframe_reporter )
		self.btn_talentdoc.Bind( wx.EVT_BUTTON, self.openframe_talentdocs )

	def __del__( self ):
		pass


	# Virtual event handlers, overide them in your derived class
	def openframe_analytics( self, event ):
		event.Skip()

	def openframe_splanner( self, event ):
		event.Skip()

	def openframe_reporter( self, event ):
		event.Skip()

	def openframe_talentdocs( self, event ):
		event.Skip()


###########################################################################
## Class analyticsFrame
###########################################################################

class analyticsFrame ( wx.Frame ):

	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"TM: Talent Analytics", pos = wx.DefaultPosition, size = wx.Size( 500,448 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
		self.SetFont( wx.Font( 9, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Museo 500" ) )
		self.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )

		mainsizer = wx.BoxSizer( wx.VERTICAL )

		self.title = wx.StaticText( self, wx.ID_ANY, u"Talent Analytics", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
		self.title.Wrap( -1 )

		self.title.SetFont( wx.Font( 14, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, True, "Museo 500" ) )
		self.title.SetForegroundColour( wx.Colour( 255, 255, 255 ) )
		self.title.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )

		mainsizer.Add( self.title, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		ui_sizer = wx.BoxSizer( wx.HORIZONTAL )

		input_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Input" ), wx.VERTICAL )

		self.txt_zhpla = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"1. ZHPLA file (.xlsx)", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_zhpla.Wrap( -1 )

		input_sizer.Add( self.txt_zhpla, 0, wx.ALL, 5 )

		self.zhpla_picker = wx.FilePickerCtrl( input_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
		input_sizer.Add( self.zhpla_picker, 0, wx.ALL, 5 )

		self.txt_zpdev = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"2. ZPDEV file (.xlsx)", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_zpdev.Wrap( -1 )

		input_sizer.Add( self.txt_zpdev, 0, wx.ALL, 5 )

		self.zpdev_picker = wx.FilePickerCtrl( input_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
		input_sizer.Add( self.zpdev_picker, 0, wx.ALL, 5 )

		self.txt_zpdevretire = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"2. ZPDEV Retire file (.xlsx)", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_zpdevretire.Wrap( -1 )

		input_sizer.Add( self.txt_zpdevretire, 0, wx.ALL, 5 )

		self.zpdevretire_picker = wx.FilePickerCtrl( input_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
		input_sizer.Add( self.zpdevretire_picker, 0, wx.ALL, 5 )


		ui_sizer.Add( input_sizer, 1, wx.EXPAND, 5 )

		process_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Process" ), wx.VERTICAL )

		self.rdbtn_cleanonly = wx.RadioButton( process_sizer.GetStaticBox(), wx.ID_ANY, u"Clean Only", wx.DefaultPosition, wx.DefaultSize, 0 )
		process_sizer.Add( self.rdbtn_cleanonly, 0, wx.ALL, 5 )

		self.rdbtn_combine = wx.RadioButton( process_sizer.GetStaticBox(), wx.ID_ANY, u"Combine Only", wx.DefaultPosition, wx.DefaultSize, 0 )
		process_sizer.Add( self.rdbtn_combine, 0, wx.ALL, 5 )

		self.rdbtn_cleancombine = wx.RadioButton( process_sizer.GetStaticBox(), wx.ID_ANY, u"Clean + Combine", wx.DefaultPosition, wx.DefaultSize, 0 )
		process_sizer.Add( self.rdbtn_cleancombine, 0, wx.ALL, 5 )

		self.chbox_autoopen = wx.CheckBox( process_sizer.GetStaticBox(), wx.ID_ANY, u"Open when finished", wx.DefaultPosition, wx.DefaultSize, 0 )
		process_sizer.Add( self.chbox_autoopen, 0, wx.ALL, 5 )

		self.btn_process = wx.Button( process_sizer.GetStaticBox(), wx.ID_ANY, u"Process Data", wx.DefaultPosition, wx.DefaultSize, 0 )
		process_sizer.Add( self.btn_process, 0, wx.ALIGN_CENTER|wx.ALL, 5 )


		ui_sizer.Add( process_sizer, 1, wx.EXPAND, 5 )

		output_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Output" ), wx.VERTICAL )

		self.btn_openfolder = wx.Button( output_sizer.GetStaticBox(), wx.ID_ANY, u"Open Folder", wx.DefaultPosition, wx.DefaultSize, 0 )
		output_sizer.Add( self.btn_openfolder, 0, wx.ALL, 5 )

		self.btn_openzhpla = wx.Button( output_sizer.GetStaticBox(), wx.ID_ANY, u"ZHPLA", wx.DefaultPosition, wx.DefaultSize, 0 )
		output_sizer.Add( self.btn_openzhpla, 0, wx.ALL, 5 )

		self.btn_openzpdev = wx.Button( output_sizer.GetStaticBox(), wx.ID_ANY, u"ZPDEV", wx.DefaultPosition, wx.DefaultSize, 0 )
		output_sizer.Add( self.btn_openzpdev, 0, wx.ALL, 5 )

		self.btn_openanalytics = wx.Button( output_sizer.GetStaticBox(), wx.ID_ANY, u"Talent Analytics", wx.DefaultPosition, wx.DefaultSize, 0 )
		output_sizer.Add( self.btn_openanalytics, 0, wx.ALL, 5 )


		ui_sizer.Add( output_sizer, 1, wx.EXPAND, 5 )


		mainsizer.Add( ui_sizer, 1, wx.EXPAND, 5 )

		self.txt_message = wx.StaticText( self, wx.ID_ANY, u"Welcome to TM Analytics", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_message.Wrap( -1 )

		self.txt_message.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )

		mainsizer.Add( self.txt_message, 0, wx.ALL, 5 )


		self.SetSizer( mainsizer )
		self.Layout()

		self.Centre( wx.BOTH )

		# Connect Events
		self.btn_process.Bind( wx.EVT_BUTTON, self.processData )
		self.btn_openfolder.Bind( wx.EVT_BUTTON, self.openFolder )
		self.btn_openzhpla.Bind( wx.EVT_BUTTON, self.openzhpla )
		self.btn_openzpdev.Bind( wx.EVT_BUTTON, self.openzpdev )
		self.btn_openanalytics.Bind( wx.EVT_BUTTON, self.openanalytics )

	def __del__( self ):
		pass


	# Virtual event handlers, overide them in your derived class
	def processData( self, event ):
		event.Skip()

	def openFolder( self, event ):
		event.Skip()

	def openzhpla( self, event ):
		event.Skip()

	def openzpdev( self, event ):
		event.Skip()

	def openanalytics( self, event ):
		event.Skip()


###########################################################################
## Class splannerFrame
###########################################################################

class splannerFrame ( wx.Frame ):

	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"TM: Succession Planner", pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
		self.SetFont( wx.Font( 9, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Museo 500" ) )
		self.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )

		mainsizer = wx.BoxSizer( wx.VERTICAL )

		self.title = wx.StaticText( self, wx.ID_ANY, u"Succession Planner", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
		self.title.Wrap( -1 )

		self.title.SetFont( wx.Font( 14, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, True, "Museo 500" ) )
		self.title.SetForegroundColour( wx.Colour( 255, 255, 255 ) )
		self.title.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )

		mainsizer.Add( self.title, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		ui_sizer = wx.BoxSizer( wx.HORIZONTAL )

		input_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Input" ), wx.VERTICAL )

		self.txt_selectfile = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"Select sp file", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_selectfile.Wrap( -1 )

		input_sizer.Add( self.txt_selectfile, 0, wx.ALL, 5 )

		self.sp_picker = wx.FilePickerCtrl( input_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
		input_sizer.Add( self.sp_picker, 0, wx.ALL, 5 )

		self.txt_selectmediafolder = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"Select media folder (optional)", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_selectmediafolder.Wrap( -1 )

		input_sizer.Add( self.txt_selectmediafolder, 0, wx.ALL, 5 )

		self.mediafolder_picker = wx.DirPickerCtrl( input_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a folder", wx.DefaultPosition, wx.DefaultSize, wx.DIRP_DEFAULT_STYLE )
		input_sizer.Add( self.mediafolder_picker, 0, wx.ALL, 5 )

		self.chbox_autoopen = wx.CheckBox( input_sizer.GetStaticBox(), wx.ID_ANY, u"Open when finished", wx.DefaultPosition, wx.DefaultSize, 0 )
		input_sizer.Add( self.chbox_autoopen, 0, wx.ALL, 5 )

		self.btn_process = wx.Button( input_sizer.GetStaticBox(), wx.ID_ANY, u"Generate SP", wx.DefaultPosition, wx.DefaultSize, 0 )
		input_sizer.Add( self.btn_process, 0, wx.ALL, 5 )


		ui_sizer.Add( input_sizer, 1, wx.EXPAND, 5 )

		output_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Output" ), wx.VERTICAL )

		self.btn_openfolder = wx.Button( output_sizer.GetStaticBox(), wx.ID_ANY, u"Open Folder", wx.DefaultPosition, wx.DefaultSize, 0 )
		output_sizer.Add( self.btn_openfolder, 0, wx.ALIGN_CENTER|wx.ALL, 5 )


		ui_sizer.Add( output_sizer, 1, wx.EXPAND, 5 )


		mainsizer.Add( ui_sizer, 1, wx.EXPAND, 5 )

		self.txt_message = wx.StaticText( self, wx.ID_ANY, u"Welcome to Succession Planner", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_message.Wrap( -1 )

		self.txt_message.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )

		mainsizer.Add( self.txt_message, 0, wx.ALL, 5 )


		self.SetSizer( mainsizer )
		self.Layout()

		self.Centre( wx.BOTH )

		# Connect Events
		self.btn_process.Bind( wx.EVT_BUTTON, self.process_sp )
		self.btn_openfolder.Bind( wx.EVT_BUTTON, self.openFolder )

	def __del__( self ):
		pass


	# Virtual event handlers, overide them in your derived class
	def process_sp( self, event ):
		event.Skip()

	def openFolder( self, event ):
		event.Skip()


###########################################################################
## Class reporterFrame
###########################################################################

class reporterFrame ( wx.Frame ):

	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"TM: Talent Reporter", pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
		self.SetFont( wx.Font( 9, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Museo 500" ) )
		self.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )

		mainsizer = wx.BoxSizer( wx.VERTICAL )

		self.title = wx.StaticText( self, wx.ID_ANY, u"Talent Reporter", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
		self.title.Wrap( -1 )

		self.title.SetFont( wx.Font( 14, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, True, "Museo 500" ) )
		self.title.SetForegroundColour( wx.Colour( 255, 255, 255 ) )
		self.title.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )

		mainsizer.Add( self.title, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		ui_sizer = wx.BoxSizer( wx.HORIZONTAL )

		input_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Source file" ), wx.VERTICAL )

		self.txt_selectfile = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"Select file containing ID (.xlsx)", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_selectfile.Wrap( -1 )

		input_sizer.Add( self.txt_selectfile, 0, wx.ALL, 5 )

		self.report_picker = wx.FilePickerCtrl( input_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
		input_sizer.Add( self.report_picker, 0, wx.ALL, 5 )

		self.txt_filename = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"filename.pdf", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_filename.Wrap( -1 )

		input_sizer.Add( self.txt_filename, 0, wx.ALL, 5 )

		self.txt_selectreporttype = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"Select report type:", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_selectreporttype.Wrap( -1 )

		input_sizer.Add( self.txt_selectreporttype, 0, wx.ALL, 5 )

		choice_pdftypeChoices = [ u"BePCB", u"CV", u"LC" ]
		self.choice_pdftype = wx.Choice( input_sizer.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, choice_pdftypeChoices, 0 )
		self.choice_pdftype.SetSelection( 0 )
		input_sizer.Add( self.choice_pdftype, 0, wx.ALL|wx.EXPAND, 5 )


		ui_sizer.Add( input_sizer, 1, wx.EXPAND, 5 )

		output_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Output" ), wx.VERTICAL )

		self.chbox_autoopen = wx.CheckBox( output_sizer.GetStaticBox(), wx.ID_ANY, u"Auto Open", wx.DefaultPosition, wx.DefaultSize, 0 )
		output_sizer.Add( self.chbox_autoopen, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		self.btn_split = wx.Button( output_sizer.GetStaticBox(), wx.ID_ANY, u"Split", wx.DefaultPosition, wx.DefaultSize, 0 )
		output_sizer.Add( self.btn_split, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		self.txt_reportcount = wx.StaticText( output_sizer.GetStaticBox(), wx.ID_ANY, u"Report count:", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_reportcount.Wrap( -1 )

		output_sizer.Add( self.txt_reportcount, 0, wx.ALL, 5 )

		self.btn_openfolder = wx.Button( output_sizer.GetStaticBox(), wx.ID_ANY, u"Open folder", wx.DefaultPosition, wx.DefaultSize, 0 )
		output_sizer.Add( self.btn_openfolder, 0, wx.ALIGN_CENTER|wx.ALL, 5 )


		ui_sizer.Add( output_sizer, 1, wx.EXPAND, 5 )


		mainsizer.Add( ui_sizer, 1, wx.ALL|wx.EXPAND, 5 )

		self.txt_message = wx.StaticText( self, wx.ID_ANY, u"Welcome to Reporter", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_message.Wrap( -1 )

		self.txt_message.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )

		mainsizer.Add( self.txt_message, 0, wx.ALL, 5 )


		self.SetSizer( mainsizer )
		self.Layout()

		self.Centre( wx.BOTH )

		# Connect Events
		self.report_picker.Bind( wx.EVT_FILEPICKER_CHANGED, self.changefilename )
		self.btn_split.Bind( wx.EVT_BUTTON, self.splitReport )
		self.btn_openfolder.Bind( wx.EVT_BUTTON, self.openFolder )

	def __del__( self ):
		pass


	# Virtual event handlers, overide them in your derived class
	def changefilename( self, event ):
		event.Skip()

	def splitReport( self, event ):
		event.Skip()

	def openFolder( self, event ):
		event.Skip()


###########################################################################
## Class talentdocsFrame
###########################################################################

class talentdocsFrame ( wx.Frame ):

	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"TM: TalentDocs", pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
		self.SetFont( wx.Font( 9, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Museo 500" ) )
		self.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )

		mainsizer = wx.BoxSizer( wx.VERTICAL )

		self.title = wx.StaticText( self, wx.ID_ANY, u"TalentDocs", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
		self.title.Wrap( -1 )

		self.title.SetFont( wx.Font( 14, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, True, "Museo 500" ) )
		self.title.SetForegroundColour( wx.Colour( 255, 255, 255 ) )
		self.title.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )

		mainsizer.Add( self.title, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		ui_sizer = wx.BoxSizer( wx.HORIZONTAL )

		input_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Input" ), wx.VERTICAL )

		self.txt_filepicker = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"Select file (.xlsx)", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_filepicker.Wrap( -1 )

		input_sizer.Add( self.txt_filepicker, 0, wx.ALL, 5 )

		self.file_picker = wx.FilePickerCtrl( input_sizer.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
		input_sizer.Add( self.file_picker, 0, wx.ALL, 5 )

		self.txt_filename = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"filename.xlsx", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_filename.Wrap( -1 )

		input_sizer.Add( self.txt_filename, 0, wx.ALL, 5 )

		selectsheet_sizer = wx.BoxSizer( wx.HORIZONTAL )

		sheet_sizer = wx.BoxSizer( wx.VERTICAL )

		self.txt_selectsheet = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"Select sheet", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_selectsheet.Wrap( -1 )

		sheet_sizer.Add( self.txt_selectsheet, 0, wx.ALL, 5 )

		choice_sheetChoices = []
		self.choice_sheet = wx.Choice( input_sizer.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, choice_sheetChoices, 0 )
		self.choice_sheet.SetSelection( 0 )
		sheet_sizer.Add( self.choice_sheet, 0, wx.ALL, 5 )


		selectsheet_sizer.Add( sheet_sizer, 1, wx.EXPAND, 5 )

		column_sizer = wx.BoxSizer( wx.VERTICAL )

		self.txt_selectcolumn = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"Select ID column", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_selectcolumn.Wrap( -1 )

		column_sizer.Add( self.txt_selectcolumn, 0, wx.ALL, 5 )

		choice_columnChoices = []
		self.choice_column = wx.Choice( input_sizer.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, choice_columnChoices, 0 )
		self.choice_column.SetSelection( 0 )
		column_sizer.Add( self.choice_column, 0, wx.ALL, 5 )


		selectsheet_sizer.Add( column_sizer, 1, wx.EXPAND, 5 )


		input_sizer.Add( selectsheet_sizer, 1, wx.EXPAND, 5 )

		self.txt_idcount = wx.StaticText( input_sizer.GetStaticBox(), wx.ID_ANY, u"ID Count: ", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_idcount.Wrap( -1 )

		self.txt_idcount.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWTEXT ) )

		input_sizer.Add( self.txt_idcount, 0, wx.ALL, 5 )


		ui_sizer.Add( input_sizer, 1, wx.EXPAND, 5 )

		process_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Process" ), wx.VERTICAL )

		self.txt_selectfile = wx.StaticText( process_sizer.GetStaticBox(), wx.ID_ANY, u"Select file to download:", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_selectfile.Wrap( -1 )

		process_sizer.Add( self.txt_selectfile, 0, wx.ALL, 5 )

		choice_filetypeChoices = [ u"BePCB", u"CV", u"DC", u"DW", u"LC", u"Profile Picture" ]
		self.choice_filetype = wx.Choice( process_sizer.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, choice_filetypeChoices, 0 )
		self.choice_filetype.SetSelection( 0 )
		process_sizer.Add( self.choice_filetype, 0, wx.ALL|wx.EXPAND, 5 )

		self.btn_process = wx.Button( process_sizer.GetStaticBox(), wx.ID_ANY, u"Process", wx.DefaultPosition, wx.DefaultSize, 0 )
		process_sizer.Add( self.btn_process, 0, wx.ALIGN_CENTER|wx.ALL, 5 )


		ui_sizer.Add( process_sizer, 1, wx.EXPAND, 5 )

		output_sizer = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"Output" ), wx.VERTICAL )

		self.btn_openfolder = wx.Button( output_sizer.GetStaticBox(), wx.ID_ANY, u"open folder", wx.DefaultPosition, wx.DefaultSize, 0 )
		output_sizer.Add( self.btn_openfolder, 0, wx.ALIGN_CENTER|wx.ALL, 5 )


		ui_sizer.Add( output_sizer, 1, wx.EXPAND, 5 )


		mainsizer.Add( ui_sizer, 1, wx.EXPAND, 5 )

		self.txt_message = wx.StaticText( self, wx.ID_ANY, u"Welcome to TalentDocs", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_message.Wrap( -1 )

		self.txt_message.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )

		mainsizer.Add( self.txt_message, 0, wx.ALL, 5 )


		self.SetSizer( mainsizer )
		self.Layout()

		self.Centre( wx.BOTH )

		# Connect Events
		self.file_picker.Bind( wx.EVT_FILEPICKER_CHANGED, self.onFileChanged )
		self.choice_sheet.Bind( wx.EVT_CHOICE, self.showColumns )
		self.choice_column.Bind( wx.EVT_CHOICE, self.setIdCount )
		self.btn_process.Bind( wx.EVT_BUTTON, self.downloadFiles )
		self.btn_openfolder.Bind( wx.EVT_BUTTON, self.openFolder )

	def __del__( self ):
		pass


	# Virtual event handlers, overide them in your derived class
	def onFileChanged( self, event ):
		event.Skip()

	def showColumns( self, event ):
		event.Skip()

	def setIdCount( self, event ):
		event.Skip()

	def downloadFiles( self, event ):
		event.Skip()

	def openFolder( self, event ):
		event.Skip()


