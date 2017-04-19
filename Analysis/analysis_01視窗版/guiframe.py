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
## Class baseFormWind
###########################################################################

class baseFormWind ( wx.Frame ):
	
	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"件號比對器", pos = wx.DefaultPosition, size = wx.Size( 233,235 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_MENU ) )
		
		bSizer1 = wx.BoxSizer( wx.VERTICAL )
		
		gSizer2 = wx.GridSizer( 0, 2, 0, 0 )
		
		self.staticText2 = wx.StaticText( self, wx.ID_ANY, u"First sheet", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.staticText2.Wrap( -1 )
		gSizer2.Add( self.staticText2, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5 )
		
		self.m_textCtrl1 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		gSizer2.Add( self.m_textCtrl1, 0, wx.ALIGN_RIGHT|wx.ALL, 5 )
		
		self.staticText4 = wx.StaticText( self, wx.ID_ANY, u"Second sheet", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.staticText4.Wrap( -1 )
		gSizer2.Add( self.staticText4, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL, 5 )
		
		self.m_textCtrl2 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		gSizer2.Add( self.m_textCtrl2, 0, wx.ALIGN_RIGHT|wx.ALL, 5 )
		
		
		bSizer1.Add( gSizer2, 1, 0, 5 )
		
		sbSizer4 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, wx.EmptyString ), wx.VERTICAL )
		
		self.m_button1 = wx.Button( sbSizer4.GetStaticBox(), wx.ID_ANY, u"開始分析", wx.DefaultPosition, wx.DefaultSize, 0 )
		sbSizer4.Add( self.m_button1, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 5 )
		
		
		bSizer1.Add( sbSizer4, 1, wx.EXPAND, 5 )
		
		sbSizer3 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, u"狀態列" ), wx.VERTICAL )
		
		self.m_staticText4 = wx.StaticText( sbSizer3.GetStaticBox(), wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText4.Wrap( -1 )
		sbSizer3.Add( self.m_staticText4, 0, wx.ALL, 5 )
		
		
		bSizer1.Add( sbSizer3, 1, wx.EXPAND, 5 )
		
		
		self.SetSizer( bSizer1 )
		self.Layout()
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.m_button1.Bind( wx.EVT_BUTTON, self.analysis )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def analysis( self, event ):
		event.Skip()
	

