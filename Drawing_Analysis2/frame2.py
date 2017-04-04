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
## Class MyFrame1
###########################################################################

class MyFrame1 ( wx.Frame ):
	
	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"圖譜分析器 v2.0", pos = wx.DefaultPosition, size = wx.Size( 460,560 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_MENU ) )
		
		bSizer1 = wx.BoxSizer( wx.VERTICAL )
		
		sbSizer41 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"步驟1: 載入圖譜總表" ), wx.VERTICAL )
		
		gSizer2 = wx.GridSizer( 0, 2, 0, 0 )
		
		self.m_textCtrl1 = wx.TextCtrl( sbSizer41.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 250,-1 ), 0 )
		gSizer2.Add( self.m_textCtrl1, 0, wx.ALL, 5 )
		
		self.m_button3 = wx.Button( sbSizer41.GetStaticBox(), wx.ID_ANY, u"指定總表檔案位置", wx.DefaultPosition, wx.DefaultSize, 0 )
		gSizer2.Add( self.m_button3, 0, wx.ALIGN_RIGHT|wx.ALL, 5 )
		
		self.m_button1 = wx.Button( sbSizer41.GetStaticBox(), wx.ID_ANY, u"載入圖譜", wx.DefaultPosition, wx.DefaultSize, 0 )
		gSizer2.Add( self.m_button1, 0, wx.ALL, 5 )
		
		
		sbSizer41.Add( gSizer2, 1, wx.EXPAND, 5 )
		
		
		bSizer1.Add( sbSizer41, 1, wx.BOTTOM|wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP, 8 )
		
		sbSizer6 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"步驟2: 載入各件號分頁檔案" ), wx.HORIZONTAL )
		
		gSizer3 = wx.GridSizer( 0, 2, 0, 0 )
		
		self.m_textCtrl3 = wx.TextCtrl( sbSizer6.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 250,-1 ), 0 )
		gSizer3.Add( self.m_textCtrl3, 0, wx.ALL, 5 )
		
		self.m_button4 = wx.Button( sbSizer6.GetStaticBox(), wx.ID_ANY, u"指定分頁檔案位置", wx.DefaultPosition, wx.DefaultSize, 0 )
		gSizer3.Add( self.m_button4, 0, wx.ALIGN_RIGHT|wx.ALL, 5 )
		
		self.m_button5 = wx.Button( sbSizer6.GetStaticBox(), wx.ID_ANY, u"載入分頁檔案", wx.DefaultPosition, wx.DefaultSize, 0 )
		gSizer3.Add( self.m_button5, 0, wx.ALL, 5 )
		
		
		sbSizer6.Add( gSizer3, 1, wx.EXPAND, 5 )
		
		
		bSizer1.Add( sbSizer6, 1, wx.BOTTOM|wx.EXPAND|wx.FIXED_MINSIZE|wx.LEFT|wx.RIGHT|wx.TOP, 8 )
		
		sbSizer4 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"步驟3: 展開分析" ), wx.VERTICAL )
		
		self.m_button2 = wx.Button( sbSizer4.GetStaticBox(), wx.ID_ANY, u"開始分析", wx.DefaultPosition, wx.DefaultSize, 0 )
		sbSizer4.Add( self.m_button2, 0, wx.ALIGN_CENTER|wx.ALL, 5 )
		
		self.m_staticText1 = wx.StaticText( sbSizer4.GetStaticBox(), wx.ID_ANY, u"狀態訊息", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText1.Wrap( -1 )
		sbSizer4.Add( self.m_staticText1, 0, wx.ALL, 5 )
		
		self.m_textCtrl2 = wx.TextCtrl( sbSizer4.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 450,250 ), wx.TE_MULTILINE )
		sbSizer4.Add( self.m_textCtrl2, 0, wx.ALL|wx.EXPAND, 5 )
		
		
		bSizer1.Add( sbSizer4, 2, wx.BOTTOM|wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP, 8 )
		
		
		self.SetSizer( bSizer1 )
		self.Layout()
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.m_button3.Bind( wx.EVT_BUTTON, self.openfile )
		self.m_button1.Bind( wx.EVT_BUTTON, self.loadExcel )
		self.m_button4.Bind( wx.EVT_BUTTON, self.OpenExcelSheet )
		self.m_button5.Bind( wx.EVT_BUTTON, self.loadExcelSheet )
		self.m_button2.Bind( wx.EVT_BUTTON, self.analysis )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def openfile( self, event ):
		event.Skip()
	
	def loadExcel( self, event ):
		event.Skip()
	
	def OpenExcelSheet( self, event ):
		event.Skip()
	
	def loadExcelSheet( self, event ):
		event.Skip()
	
	def analysis( self, event ):
		event.Skip()
	

