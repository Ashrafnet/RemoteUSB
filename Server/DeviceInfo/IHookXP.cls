VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IHookXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2009 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Implement this interface in objects that sink messages
' using the subclassing technique offered by MHookXP.
Public Function Message(ByVal hWnd As Long, _
                        ByVal uiMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long, _
                        ByVal dwRefData As Long) As Long
End Function

' *** Prototype:
'Private Function IHookXP_Message(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
'   Debug.Print "hWnd: 0x"; Hex$(hWnd), "Msg: 0x"; Hex$(uiMsg), _
'               "wParam: 0x"; Hex$(wParam), "lParam: 0x"; Hex$(lParam), _
'               "RefData: "; dwRefData
'   IHookXP_Message = HookDefault(hWnd, uiMsg, wParam, lParam)
'End Function
' **************

