Attribute VB_Name = "MSysInfo"
' *************************************************************************
'  Copyright ©2009 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Win32 API declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassname Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Constants used with above APIs
Private Const GWL_HWNDPARENT As Long = (-8)

' *********************************************
'  Public Methods
' *********************************************
Public Function FindHiddenTopWindow() As Long
   ' This function returns the hidden toplevel window
   ' associated with the current thread of execution.
   Call EnumThreadWindows(App.ThreadID, AddressOf EnumThreadWndProc, VarPtr(FindHiddenTopWindow))
End Function

' *********************************************
'  Private Methods
' *********************************************
Private Function EnumThreadWndProc(ByVal hWnd As Long, ByVal lpResult As Long) As Long
   Dim nStyle As Long
   Dim Class As String
   
   ' Assume we will continue enumeration.
   EnumThreadWndProc = True
   
   ' Test to see if this window is parented.
   ' If not, it may be what we're looking for!
   If GetWindowLong(hWnd, GWL_HWNDPARENT) = 0 Then
      ' This rules out IDE windows when not compiled.
      Class = Classname(hWnd)
      ' Version agnostic test.
      If InStr(Class, "Thunder") = 1 Then
         If InStr(Class, "Main") = (Len(Class) - 3) Then
            ' Copy hWnd to result variable pointer,
            Call CopyMemory(ByVal lpResult, hWnd, 4&)
            ' and stop enumeration.
            EnumThreadWndProc = False
            #If Debugging Then
               Debug.Print Class; " hWnd=&h"; Hex$(hWnd)
               Print #hLog, Class; " hWnd=&h"; Hex$(hWnd)
            #End If
         End If
      End If
   End If
End Function

Private Function Classname(ByVal hWnd As Long) As String
   Dim nRet As Long
   Dim Class As String
   Const MaxLen As Long = 256
   
   ' Retrieve classname of passed window.
   Class = String$(MaxLen, 0)
   nRet = GetClassname(hWnd, Class, MaxLen)
   If nRet Then Classname = Left$(Class, nRet)
End Function


