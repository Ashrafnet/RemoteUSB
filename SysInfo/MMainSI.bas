Attribute VB_Name = "MMainSI"
' *************************************************************************
'  Copyright ©2009 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Global variables
Public g_SysInfoGeneral As CSysInfoGeneral
Public g_SysInfoDevice As CSysInfoDevice
Public g_SysInfoPower As CSysInfoPower

Public hLog As Long

Public Sub Main()
   Dim frm As FSysInfo
   ' Create logfile
   #If Debugging Then
      hLog = FreeFile
      Open App.Path & "\" & App.Title & ".log" For Output As #hLog
   #End If
   ' Create global instance of SysInfo class.
   Set g_SysInfoDevice = New CSysInfoDevice
   Set g_SysInfoGeneral = New CSysInfoGeneral
   Set g_SysInfoPower = New CSysInfoPower
   ' Fire up main application form.
   Set frm = New FSysInfo
   frm.Show
End Sub


