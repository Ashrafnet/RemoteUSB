Attribute VB_Name = "modDevice"
Declare Sub RegisterDeviceNotification Lib "user32.dll" _
Alias "RegisterDeviceNotificationA" ( _
ByVal hRecipient As Long, _
NotificationFilter As Any, _
ByVal Flags As Long)

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_DRAWCLIPBOARD = &H308
Public Const GWL_WNDPROC = (-4)



Public Function MyWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Debug.Print "WM_DEVICECHANGE " & WM_DEVICECHANGE
    
    If Msg = WM_DEVICECHANGE Then
        Debug.Print "wParam " & wParam
        
        Select Case wParam
            Case &H8000& ' Device Attached
           frmMain.DeviceChanging
            
            Case &H8004& ' Device Removed
           frmMain.DeviceChanging
            
            
            
        End Select
        MyWindowProc = 0
        Exit Function
    End If

MyWindowProc = CallWindowProc(glngPrevWndProc, hwnd, Msg, wParam, lParam)
End Function
