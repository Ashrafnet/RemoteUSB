Attribute VB_Name = "modWindow"
Option Explicit
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long


' Threads
Public Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CreateEvent& Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpname As String)

Public lThreadHandle1 As Long
Public lThreadHandle2 As Long
Public lEventHandle As Long
Public MainRegWindow As Long
Public SubMainRegWindow As Long


' Enumerates all windows active on the
' the system until the last window has
' been reached.
Public Declare Function EnumWindows _
    Lib "user32" ( _
        ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) _
        As Long



' Returns the class name form which the
' the specified window was created from.
Public Declare Function GetClassName _
    Lib "user32" _
    Alias "GetClassNameA" ( _
        ByVal hwnd As Long, _
        ByVal lpClassName As String, _
        ByVal nMaxCount As Long) _
        As Long
        
' Determines if the specified window
' is minimized (iconic)
Public Declare Function IsIconic _
    Lib "user32" ( _
        ByVal hwnd As Long) _
        As Long

' Determines if the specified window
' handle is a valid handle.
Public Declare Function IsWindow _
    Lib "user32" ( _
        ByVal hwnd As Long) _
        As Long

' Determines if a window is visible on
' the system.
Public Declare Function IsWindowVisible _
    Lib "user32" ( _
        ByVal hwnd As Long) _
        As Long

' Posts a specified windows message into
' a window procedures message queue to be
' processed
Public Declare Function PostMessage _
    Lib "user32" _
    Alias "PostMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) _
        As Long

' Shows the specified window in a
' modal state specified by 'nCmdShow'
Public Declare Function ShowWindow _
    Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nCmdShow As Long) _
        As Long

Public Const SW_NORMAL      As Long = 1             ' Shows the window in its normal state
Public Const SW_MAXIMIZE    As Long = 3             ' Shows the window in its maxmized state
Public Const SW_MINIMIZE    As Long = 6             ' Sets the windows state to minimized

Public Const WM_CLOSE       As Long = &H10          ' [Windows Message] Closes the target window

Public lvwListItem          As ListItem             ' ListView List Item object

Public Function EnumVisibleWindows(ByVal hwnd As Long, ByVal lParam As Long) As Long
    
    Dim lngReturn           As Long                     ' Return Value variable
    Dim strWindowName       As String * 256             ' Holds the window caption (256 Characters max)
    Dim strWindowClass      As String * 256             ' Holds the window class name (256 Characters max)
    Dim blnIsVisible        As Boolean                  ' [Flag] Determines if a window is visible or not
    
    ' Check if the window is visible
    blnIsVisible = IsWindowVisible(hwnd)
    ' If it is then get the other information needed
    ' (Class Name & Window Caption)
    If blnIsVisible = True Then
        lngReturn = GetWindowText(hwnd, strWindowName, 256)
        If lngReturn Then
            Call GetClassName(hwnd, strWindowClass, 256)
            If InStr(1, LCase(strWindowName)) = "network connections" Then
                
                ShowWindow hwnd, 1
                
                Exit Function
            End If
            
        End If
    End If
    ' Continue the Enumeration
    EnumVisibleWindows = True
End Function

Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim sSave As String
    'Get the windowtext length
    sSave = Space$(GetWindowTextLength(hwnd) + 1)
    'get the window text
    GetWindowText hwnd, sSave, Len(sSave)
    'remove the last Chr$(0)
    sSave = Left$(sSave, Len(sSave) - 1)
    If sSave <> "" Then
        
        EnumChildProc = False
    End If
    'continue enumeration
    EnumChildProc = 1
End Function




