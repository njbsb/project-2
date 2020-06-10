# -*- coding: utf-8 -*- 

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc

###########################################################################
## Class fr_downloader
###########################################################################

class fr_downloader ( wx.Frame ):
	
	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"TTM: TE Batch Downloader", pos = wx.DefaultPosition, size = wx.Size( 644,350 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		self.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )
		
		main_layout = wx.BoxSizer( wx.VERTICAL )
		
		self.txt_title = wx.StaticText( self, wx.ID_ANY, u"Talent Engine Batch Downloader", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_title.Wrap( -1 )
		self.txt_title.SetFont( wx.Font( 12, 74, 90, 92, False, "Century Gothic" ) )
		self.txt_title.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		
		main_layout.Add( self.txt_title, 0, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		b_layout_h = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, wx.EmptyString ), wx.HORIZONTAL )
		
		b_layout_ID = wx.StaticBoxSizer( wx.StaticBox( b_layout_h.GetStaticBox(), wx.ID_ANY, u"ID" ), wx.VERTICAL )
		
		self.txtfld_strings = wx.TextCtrl( b_layout_ID.GetStaticBox(), wx.ID_ANY, u"input example: 123, 234, 345", wx.DefaultPosition, wx.Size( 200,100 ), 0|wx.NO_BORDER )
		self.txtfld_strings.SetFont( wx.Font( 9, 74, 93, 90, False, "Arial" ) )
		
		b_layout_ID.Add( self.txtfld_strings, 0, wx.ALIGN_CENTER|wx.ALIGN_TOP|wx.ALL, 5 )
		
		self.m_staticText5 = wx.StaticText( b_layout_ID.GetStaticBox(), wx.ID_ANY, u"or", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText5.Wrap( -1 )
		self.m_staticText5.SetFont( wx.Font( 8, 74, 93, 90, False, "Century Gothic" ) )
		
		b_layout_ID.Add( self.m_staticText5, 0, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		self.btn_selectfile = wx.Button( b_layout_ID.GetStaticBox(), wx.ID_ANY, u"Select file...", wx.DefaultPosition, wx.DefaultSize, 0|wx.NO_BORDER )
		self.btn_selectfile.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		self.btn_selectfile.SetBackgroundColour( wx.Colour( 36, 106, 115 ) )
		
		b_layout_ID.Add( self.btn_selectfile, 0, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		self.txt_filepath = wx.StaticText( b_layout_ID.GetStaticBox(), wx.ID_ANY, u"file path .txt", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_filepath.Wrap( -1 )
		b_layout_ID.Add( self.txt_filepath, 0, wx.ALL, 5 )
		
		self.txt_countID = wx.StaticText( b_layout_ID.GetStaticBox(), wx.ID_ANY, u"count", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_countID.Wrap( -1 )
		self.txt_countID.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		
		b_layout_ID.Add( self.txt_countID, 0, wx.ALL, 5 )
		
		
		b_layout_h.Add( b_layout_ID, 1, wx.EXPAND, 5 )
		
		b_layout_selection = wx.StaticBoxSizer( wx.StaticBox( b_layout_h.GetStaticBox(), wx.ID_ANY, u"Selection" ), wx.VERTICAL )
		
		dropdown_filetypeChoices = [ u"Select file", u"BePCB", u"CV", u"DC", u"DW", u"LC", u"Profile Picture" ]
		self.dropdown_filetype = wx.Choice( b_layout_selection.GetStaticBox(), wx.ID_ANY, wx.DefaultPosition, wx.Size( 180,-1 ), dropdown_filetypeChoices, 0|wx.NO_BORDER )
		self.dropdown_filetype.SetSelection( 0 )
		self.dropdown_filetype.SetForegroundColour( wx.Colour( 92, 79, 63 ) )
		self.dropdown_filetype.SetBackgroundColour( wx.Colour( 221, 216, 196 ) )
		
		b_layout_selection.Add( self.dropdown_filetype, 0, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		self.btn_download = wx.Button( b_layout_selection.GetStaticBox(), wx.ID_ANY, u"Download", wx.DefaultPosition, wx.DefaultSize, 0|wx.NO_BORDER )
		self.btn_download.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		self.btn_download.SetBackgroundColour( wx.Colour( 36, 106, 115 ) )
		
		b_layout_selection.Add( self.btn_download, 0, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		self.txt_status = wx.StaticText( b_layout_selection.GetStaticBox(), wx.ID_ANY, u"status -....-", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_status.Wrap( -1 )
		self.txt_status.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		
		b_layout_selection.Add( self.txt_status, 0, wx.ALL, 5 )
		
		self.txt_countRemaining = wx.StaticText( b_layout_selection.GetStaticBox(), wx.ID_ANY, u"count_remaining", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_countRemaining.Wrap( -1 )
		self.txt_countRemaining.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		
		b_layout_selection.Add( self.txt_countRemaining, 0, wx.ALL, 5 )
		
		self.txtfld_failed = wx.TextCtrl( b_layout_selection.GetStaticBox(), wx.ID_ANY, u"123,123,123,123,123,", wx.DefaultPosition, wx.Size( 200,100 ), 0|wx.NO_BORDER )
		self.txtfld_failed.SetForegroundColour( wx.Colour( 36, 106, 115 ) )
		self.txtfld_failed.SetBackgroundColour( wx.Colour( 0, 177, 169 ) )
		self.txtfld_failed.Enable( False )
		
		b_layout_selection.Add( self.txtfld_failed, 0, wx.ALL, 5 )
		
		
		b_layout_h.Add( b_layout_selection, 1, wx.EXPAND, 5 )
		
		b_layout_output = wx.StaticBoxSizer( wx.StaticBox( b_layout_h.GetStaticBox(), wx.ID_ANY, u"Output" ), wx.VERTICAL )
		
		self.txt_statusFinal = wx.StaticText( b_layout_output.GetStaticBox(), wx.ID_ANY, u"statusFinal", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.txt_statusFinal.Wrap( -1 )
		self.txt_statusFinal.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )
		
		b_layout_output.Add( self.txt_statusFinal, 0, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		self.btn_openoutputdir = wx.Button( b_layout_output.GetStaticBox(), wx.ID_ANY, u"Open Folder...", wx.Point( -1,-20 ), wx.DefaultSize, 0 )
		b_layout_output.Add( self.btn_openoutputdir, 0, wx.ALIGN_CENTER|wx.TOP, 5 )
		
		
		b_layout_h.Add( b_layout_output, 1, wx.EXPAND, 5 )
		
		
		main_layout.Add( b_layout_h, 1, wx.EXPAND, 5 )
		
		
		self.SetSizer( main_layout )
		self.Layout()
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.txtfld_strings.Bind( wx.EVT_TEXT, self.update_id_count )
		self.btn_selectfile.Bind( wx.EVT_BUTTON, self.selectFile )
		self.btn_download.Bind( wx.EVT_BUTTON, self.download_file )
		self.btn_openoutputdir.Bind( wx.EVT_BUTTON, self.openDir )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def update_id_count( self, event ):
		event.Skip()
	
	def selectFile( self, event ):
		event.Skip()
	
	def download_file( self, event ):
		event.Skip()
	
	def openDir( self, event ):
		event.Skip()
	

