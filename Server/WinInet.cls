VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinInet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' private Module variables
Private mlConnectionNumber As Long
Private mbDisconnectOnTerminate As Boolean

' for list dun's function
Private Type RAS_ENTRIES
dwSize As Long
szEntryname(256) As Byte
End Type
Private Declare Function RasEnumEntriesA Lib "rasapi32.dll" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long

' for activeconnection funciton
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal sSubKey As String, hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal sKeyValue As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, nSizeData As Long) As Long

' for Dial and Hangup functions
Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwnd As Long, ByVal sConnectoid As String, ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long
' Returns ERROR_SUCCESS if successfull or one of the following error codes
' ERROR_INVALID_PARAMETER - one or more parameters are incorrect
' ERROR_NO_CONNECTION - There is a problem with the dial-up connection
' ERROR_USER_DISCONNECTION - The user clicked either the work offline or cancel button on the dialog box
Private Declare Function InternetHangUp Lib "wininet.dll" (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long
' Returns ERROR_SUCCESS if successfull or an error value otherwise
' Flags for InternetAutodial
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = &H1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = &H2
Private Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = &H4
' Flags for InternetDial - must not conflict with InternetAutodial flags
' as they are valid here also.
Private Const INTERNET_DIAL_FORCE_PROMPT = &H2000
Private Const INTERNET_DIAL_SHOW_OFFLINE = &H4000
Private Const INTERNET_DIAL_UNATTENDED = &H8000

' Windows error constants used by all sub's
Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_INVALID_PARAMETER = 87&
' RAS error constants
Private Const RASBASE As Long = 600& 'not sure about this couldn't find raserror.h anywhere on MSDN so
'best-guessed the value based on return code of 631 for cancel button
Private Const ERROR_NO_CONNECTION = (RASBASE + 68&)
Private Const ERROR_USER_DISCONNECTION = (RASBASE + 31&)

' Events for this module
Public Event ConnectionMade()
Public Event ConnectionClosed()
'
'


Private Sub Class_Initialize()
mlConnectionNumber = 0&
mbDisconnectOnTerminate = False
End Sub
Private Sub Class_Terminate()
If mbDisconnectOnTerminate And mlConnectionNumber <> 0 Then
Call InternetHangUp(mlConnectionNumber, 0&)
End If
End Sub
Public Property Get Connected() As Boolean
    Connected = ActiveConnection()
End Property
Public Property Get DisconnectOnTerminate() As Boolean
    DisconnectOnTerminate = mbDisconnectOnTerminate
End Property
Public Property Let DisconnectOnTerminate(ByVal bValue As Boolean)
mbDisconnectOnTerminate = bValue
End Property
Public Function HangUp() As Long
    If mlConnectionNumber = 0 Then
    HangUp = -1 'no connection from this module
    Else
    HangUp = InternetHangUp(mlConnectionNumber, 0&)
    mlConnectionNumber = 0&
    RaiseEvent ConnectionClosed
    End If
End Function
Public Sub ListDUNs(sDunList() As String)
Dim plSize As Long
Dim plEntries As Long
Dim psConName As String
Dim plIndex As Long
Dim RAS(255) As RAS_ENTRIES

Erase sDunList()
RAS(0).dwSize = 264
plSize = 256 * RAS(0).dwSize
Call RasEnumEntriesA(vbNullString, vbNullString, RAS(0), plSize, plEntries)
plEntries = plEntries - 1
If plEntries >= 0 Then
ReDim sDunList(plEntries)
For plIndex = 0 To plEntries
psConName = StrConv(RAS(plIndex).szEntryname(), vbUnicode)
sDunList(plIndex) = Left$(psConName, InStr(psConName, vbNullChar) - 1)
Next plIndex
End If
End Sub
Public Function StartDUN(hwnd As Long, sDUN As String) As Long
Dim plResult As Long

If mlConnectionNumber <> 0 And ActiveConnection() Then
StartDUN = -1 'already issued a connection
Else
plResult = InternetDial(hwnd, sDUN, INTERNET_AUTODIAL_FORCE_UNATTENDED, mlConnectionNumber, 0&)
If plResult = ERROR_SUCCESS Then
RaiseEvent ConnectionMade
Else
mlConnectionNumber = 0 'somethings amiss, clear connection #
End If
StartDUN = plResult
End If
End Function

' Module support routines
'======================================================================================
Private Function ActiveConnection() As Boolean
    Dim hKey As Long
    Dim lpData As Long
    Dim nSizeData As Long
    
    'function checks registry for an active connection
    Const sSubKey = "SystemCurrentControlSetServicesRemoteAccess"
    Const sKeyValue = "Remote Connection"
    'default false
    ActiveConnection = False
    If RegOpenKey(HKEY_LOCAL_MACHINE, sSubKey, hKey) = ERROR_SUCCESS Then
        lpData = 0&
        nSizeData = Len(lpData)
        If RegQueryValueEx(hKey, sKeyValue, 0&, 0&, lpData, nSizeData) = ERROR_SUCCESS Then
            ActiveConnection = lpData <> 0
        End If
        Call RegCloseKey(hKey)
    End If
End Function

