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
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"圖譜分析器", pos = wx.DefaultPosition, size = wx.Size( 509,347 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_MENU ) )
		
		bSizer1 = wx.BoxSizer( wx.VERTICAL )
		
		sbSizer6 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"載入檔案" ), wx.VERTICAL )
		
		self.m_textCtrl1 = wx.TextCtrl( sbSizer6.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 300,-1 ), 0 )
		sbSizer6.Add( self.m_textCtrl1, 0, wx.ALL, 5 )
		
		self.m_button1 = wx.Button( sbSizer6.GetStaticBox(), wx.ID_ANY, u"載入圖譜", wx.DefaultPosition, wx.DefaultSize, 0 )
		sbSizer6.Add( self.m_button1, 0, wx.ALL, 5 )
		
		self.m_button2 = wx.Button( sbSizer6.GetStaticBox(), wx.ID_ANY, u"開始分析", wx.DefaultPosition, wx.DefaultSize, 0 )
		sbSizer6.Add( self.m_button2, 0, wx.ALL, 5 )
		
		
		bSizer1.Add( sbSizer6, 1, wx.EXPAND|wx.FIXED_MINSIZE, 5 )
		
		sbSizer4 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"狀態訊息" ), wx.VERTICAL )
		
		self.m_textCtrl2 = wx.TextCtrl( sbSizer4.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 450,250 ), wx.TE_MULTILINE )
		sbSizer4.Add( self.m_textCtrl2, 0, wx.ALL|wx.EXPAND, 5 )
		
		
		bSizer1.Add( sbSizer4, 2, wx.EXPAND, 5 )
		
		
		self.SetSizer( bSizer1 )
		self.Layout()
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.m_button1.Bind( wx.EVT_BUTTON, self.loadExcel )
		self.m_button2.Bind( wx.EVT_BUTTON, self.analysis )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def loadExcel( self, event ):
		event.Skip()
	
	def analysis( self, event ):
		event.Skip()
	

