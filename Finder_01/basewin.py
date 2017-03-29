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
## Class baseMainWindow
###########################################################################

class baseMainWindow ( wx.Frame ):
	
	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"接頭料號搜尋器 v1.0", pos = wx.DefaultPosition, size = wx.Size( 440,355 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		self.SetForegroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_BACKGROUND ) )
		self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INACTIVEBORDER ) )
		
		bSizer1 = wx.BoxSizer( wx.VERTICAL )
		
		sbSizer1 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, wx.EmptyString ), wx.VERTICAL )
		
		self.staticText1 = wx.StaticText( sbSizer1.GetStaticBox(), wx.ID_ANY, u"請輸入接頭原廠編號:", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.staticText1.Wrap( -1 )
		sbSizer1.Add( self.staticText1, 0, wx.ALL, 5 )
		
		self.text_main = wx.TextCtrl( sbSizer1.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.Point( -1,-1 ), wx.Size( 200,-1 ), 0 )
		sbSizer1.Add( self.text_main, 0, wx.ALL, 5 )
		
		
		bSizer1.Add( sbSizer1, 1, wx.EXPAND, 5 )
		
		sbSizer2 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, wx.EmptyString ), wx.VERTICAL )
		
		sbSizer2.SetMinSize( wx.Size( 1,90 ) ) 
		self.button3 = wx.Button( sbSizer2.GetStaticBox(), wx.ID_ANY, u"開始搜尋", wx.DefaultPosition, wx.DefaultSize, 0 )
		sbSizer2.Add( self.button3, 0, wx.ALL, 5 )
		
		self.button_main = wx.Button( sbSizer2.GetStaticBox(), wx.ID_ANY, u"清空", wx.DefaultPosition, wx.DefaultSize, 0 )
		sbSizer2.Add( self.button_main, 0, wx.ALL, 5 )
		
		
		bSizer1.Add( sbSizer2, 1, wx.ALIGN_CENTER|wx.EXPAND, 5 )
		
		sbSizer3 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"搜尋結果" ), wx.VERTICAL )
		
		self.m_textCtrl2 = wx.TextCtrl( sbSizer3.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 400,100 ), wx.TE_MULTILINE|wx.TE_READONLY )
		sbSizer3.Add( self.m_textCtrl2, 0, wx.ALL, 5 )
		
		
		bSizer1.Add( sbSizer3, 1, wx.EXPAND, 5 )
		
		
		self.SetSizer( bSizer1 )
		self.Layout()
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.button3.Bind( wx.EVT_BUTTON, self.finder )
		self.button_main.Bind( wx.EVT_BUTTON, self.main_button_click )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def finder( self, event ):
		event.Skip()
	
	def main_button_click( self, event ):
		event.Skip()
	

