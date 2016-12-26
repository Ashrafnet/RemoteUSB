VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSysInfoGeneral"
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

' Win32 API declarations.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Monitored messages.
Private Const WM_QUERYENDSESSION              As Long = &H11
Private Const WM_SYSCOLORCHANGE               As Long = &H15
Private Const WM_ENDSESSION                   As Long = &H16
Private Const WM_SETTINGCHANGE                As Long = &H1A
Private Const WM_ACTIVATEAPP                  As Long = &H1C
Private Const WM_FONTCHANGE                   As Long = &H1D
Private Const WM_TIMECHANGE                   As Long = &H1E
Private Const WM_DISPLAYCHANGE                As Long = &H7E
Private Const WM_NCDESTROY                    As Long = &H82
Private Const WM_POWERBROADCAST               As Long = &H218
Private Const WM_THEMECHANGED                 As Long = &H31A

' QueryEndSession logoff options
Private Const ENDSESSION_SHUTDOWN As Long = &H0
Private Const ENDSESSION_CLOSEAPP As Long = &H1
Private Const ENDSESSION_CRITICAL As Long = &H40000000
Private Const ENDSESSION_LOGOFF   As Long = &H80000000

Public Enum EndSessionFlags
   esShutdown = ENDSESSION_SHUTDOWN
   esCloseApp = ENDSESSION_CLOSEAPP
   esCritical = ENDSESSION_CRITICAL
   esLogoff = ENDSESSION_LOGOFF
End Enum

' Subclassing interface
Implements IHookXP

' System notification events.
Public Event ActivateApp(ByVal Active As Boolean)
Public Event DisplayChange(ByVal BitsPerPixel As Long, ByVal PixelsX As Long, ByVal PixelsY As Long)
Public Event EndSession(ByVal EndingInitiated As Boolean, ByVal Flag As Long)
Public Event FontChange()
Public Event QueryEndSession(ByVal Flag As EndSessionFlags, Cancel As Boolean)
Public Event SettingChange(ByVal Setting As String, ByVal Flag As Long)
Public Event SysColorChange()
Public Event ThemeChanged()
Public Event TimeChanged()

' Member variables
Private m_hWndHidden As Long

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   ' Hook into hidden toplevel message stream.
   m_hWndHidden = FindHiddenTopWindow()
   Call HookSet(m_hWndHidden, Me)
End Sub

Private Sub Class_Terminate()
   ' Tear down
   Call HookClear(m_hWndHidden, Me)
End Sub

' *********************************************
'  Public Methods
' *********************************************

' *********************************************
'  Private Methods
' *********************************************
Private Function TrimNull(ByVal StrZ As String) As String
   TrimNull = Left$(StrZ, InStr(StrZ & vbNullChar, vbNullChar) - 1)
End Function

' *********************************************
'  Implemented Subclassing Interface
' *********************************************
Private Function IHookXP_Message(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
   Dim Cancel As Boolean
   Dim EatIt As Boolean
   Dim msg As String
   
   ' Special processing for messages we care about.
   Select Case uiMsg
      Case WM_ACTIVATEAPP
         ' wParam indicates active status.
         RaiseEvent ActivateApp(CBool(wParam))
         
      Case WM_DISPLAYCHANGE
         ' wParam is BPP, lParam is X/Y.
         RaiseEvent DisplayChange(wParam, LoWord(lParam), HiWord(lParam))
         
      Case WM_ENDSESSION
         ' wParam confirms end, lParam provides reason flag.
         ' If the session is being ended, wParam is TRUE; the session
         ' can end any time after all applications have returned from
         ' processing this message. Otherwise, it is FALSE.
         RaiseEvent EndSession(CBool(wParam), lParam)
         
      Case WM_FONTCHANGE
         ' wParam and lParam are both reserved, unused.
         RaiseEvent FontChange
         
      Case WM_QUERYENDSESSION
         ' lParam provides flag that indicates reason for ending.
         RaiseEvent QueryEndSession(lParam, Cancel)
         ' Applications should respect the user's intentions and
         ' return TRUE. By default, the DefWindowProc function
         ' returns TRUE for this message.
         ' If shutting down would corrupt the system or media that
         ' is being burned, the application can return FALSE.
         IHookXP_Message = Abs(Not Cancel)
         EatIt = True
      
      Case WM_SYSCOLORCHANGE
         ' wParam and lParam are both reserved, unused.
         RaiseEvent SysColorChange
         
      Case WM_SETTINGCHANGE
         ' wParam is source Flag, lParam is source Name.
         ' Uses MMsgLookup.bas module for demo purposes!
         RaiseEvent SettingChange(SettingName(wParam, lParam), wParam)
         
      Case WM_THEMECHANGED
         ' wParam and lParam are both reserved, unused.
         RaiseEvent ThemeChanged
         
      Case WM_TIMECHANGE
         ' No parameters.
         RaiseEvent TimeChanged
         
      Case Else
         ' Uses MMsgLookup.bas module for demo purposes!
         #If Debugging Then
            msg = Format$(Now, "hh:mm:ss") & vbTab & _
                  Hex$(hWnd) & vbTab & _
                  GetMessageName(uiMsg) & vbTab & _
                  Hex$(wParam) & vbTab & _
                  Hex$(lParam)
            Debug.Print msg
            Print #hLog, msg
         #End If
   
   End Select
   
   ' Pass back to default message handler.
   If EatIt = False Then
      IHookXP_Message = HookDefault(hWnd, uiMsg, wParam, lParam)
   End If

End Function

Private Function SettingName(ByVal wParam As Long, ByVal lParam As Long) As String
   Dim sRet As String
   ' wParam is source Flag, lParam is source Name.
   ' Source name is Unicode on XP.  Might be ANSI on 9x???
   ' Flag is 1 for policy, # for SystemParametersInfo value.
   sRet = PointerToStringW(lParam)
   ' This is probably an SPI flag, try to dereference.
   If wParam > 1 Then
      sRet = GetParameterName(wParam)
   End If
   SettingName = sRet
End Function
