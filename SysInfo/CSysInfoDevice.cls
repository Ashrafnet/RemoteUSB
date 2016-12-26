VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSysInfoDevice"
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
Private Declare Function RegisterDeviceNotification Lib "user32" Alias "RegisterDeviceNotificationA" (ByVal hRecipient As Long, NotificationFilter As Any, ByVal Flags As Long) As Long
Private Declare Function UnregisterDeviceNotification Lib "user32" (ByVal hRecipient As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Monitored messages.
Private Const WM_DEVICECHANGE                 As Long = &H219

' Return this value to deny a query.
Private Const BROADCAST_QUERY_DENY            As Long = &H424D5144

' WM_DEVICECHANGE events.
Private Const DBT_DEVICEARRIVAL               As Long = &H8000&  ' system detected a new device
Private Const DBT_DEVICEQUERYREMOVE           As Long = &H8001&  ' wants to remove, may fail
Private Const DBT_DEVICEQUERYREMOVEFAILED     As Long = &H8002&  ' removal aborted
Private Const DBT_DEVICEREMOVEPENDING         As Long = &H8003&  ' about to remove, still avail.
Private Const DBT_DEVICEREMOVECOMPLETE        As Long = &H8004&  ' device is gone
Private Const DBT_DEVICETYPESPECIFIC          As Long = &H8005&  ' type specific event
Private Const DBT_CUSTOMEVENT                 As Long = &H8006&  ' user-defined event
Private Const DBT_DEVNODES_CHANGED            As Long = &H7
Private Const DBT_QUERYCHANGECONFIG           As Long = &H17
Private Const DBT_CONFIGCHANGED               As Long = &H18
Private Const DBT_CONFIGCHANGECANCELED        As Long = &H19
Private Const DBT_MONITORCHANGE               As Long = &H1B

' WM_DEVICECHANGE device types.
Private Const DBT_DEVTYP_OEM                  As Long = &H0     ' oem-defined device type
Private Const DBT_DEVTYP_DEVNODE              As Long = &H1     ' devnode number
Private Const DBT_DEVTYP_VOLUME               As Long = &H2     ' logical volume
Private Const DBT_DEVTYP_PORT                 As Long = &H3     ' serial, parallel
Private Const DBT_DEVTYP_NET                  As Long = &H4     ' network resource
Private Const DBT_DEVTYP_DEVICEINTERFACE      As Long = &H5     ' device interface class
Private Const DBT_DEVTYP_HANDLE               As Long = &H6     ' file system handle

Private Const DEVICE_NOTIFY_WINDOW_HANDLE As Long = &H0
Private Const DEVICE_NOTIFY_SERVICE_HANDLE As Long = &H1
Private Const DEVICE_NOTIFY_ALL_INTERFACE_CLASSES As Long = &H4

Private Type GUID
   Data1(1 To 4) As Byte
   Data2(1 To 2) As Byte
   Data3(1 To 2) As Byte
   Data4(1 To 8) As Byte
End Type

Private Type DEV_BROADCAST_HDR
   cbSize As Long
   DeviceType As Long
   Reserved As Long
End Type

Private Type DEV_BROADCAST_DEVICEINTERFACE
   cbSize As Long
   DeviceType As Long
   Reserved As Long
   ClassGuid As GUID
   'DeviceName As Long
   DeviceName As String * 1024
End Type

Private Type DEV_BROADCAST_HANDLE
   cbSize As Long
   DeviceType As Long
   Reserved As Long
   Handle As Long
   hDevNotify As Long
   EventGuid As GUID
   NameOffset As Long
   Data(0 To 1023) As Byte
End Type

Private Type DEV_BROADCAST_VOLUME
   cbSize As Long
   DeviceType As Long
   Reserved As Long
   UnitMask As Long
   Flags As Integer
End Type

' Subclassing interface
Implements IHookXP

' System notification events.
Public Event DeviceArrival(ByVal DevType As DeviceTypes)
Public Event DeviceQueryRemove(ByVal DevType As DeviceTypes, Cancel As Boolean)
Public Event DeviceQueryRemoveFailed(ByVal DevType As DeviceTypes)
Public Event DeviceRemovePending(ByVal DevType As DeviceTypes)
Public Event DeviceRemoveComplete(ByVal DevType As DeviceTypes)

' Enumerations
Public Enum DeviceTypes
   DeviceTypeOEM = DBT_DEVTYP_OEM                     ' OEM- or IHV-defined device type.
   DeviceTypeVolume = DBT_DEVTYP_VOLUME               ' Logical volume.
   DeviceTypePort = DBT_DEVTYP_PORT                   ' Port device (serial or parallel).
   DeviceTypeNetwork = DBT_DEVTYP_NET                 ' Network resource.
   DeviceTypeInterface = DBT_DEVTYP_DEVICEINTERFACE   ' Class of devices.
   DeviceTypeHandle = DBT_DEVTYP_HANDLE               ' File system handle.
End Enum

Public Enum DeviceFlags
   DeviceFlagRemovable = 0
   DeviceFlagMedia = 1       '#define DBTF_MEDIA   0x0001     // media comings and goings
   DeviceFlagNetwork = 2     '#define DBTF_NET     0x0002     // network volume
End Enum

' Member variables
Private m_hWndHidden As Long
Private m_hDevNotify As Long
Private m_lpDevInfo As Long
Private m_hDrvNotify(0 To 25) As Long

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   Dim NotificationFilter As DEV_BROADCAST_DEVICEINTERFACE
   Dim Flags As Long
   
   ' Hook into hidden toplevel message stream.
   m_hWndHidden = FindHiddenTopWindow()
   Call HookSet(m_hWndHidden, Me)
   
   ' Hook into notifications for device interfaces.
   NotificationFilter.cbSize = Len(NotificationFilter)
   NotificationFilter.DeviceType = DBT_DEVTYP_DEVICEINTERFACE
   Flags = DEVICE_NOTIFY_WINDOW_HANDLE Or DEVICE_NOTIFY_ALL_INTERFACE_CLASSES
   m_hDevNotify = RegisterDeviceNotification(m_hWndHidden, NotificationFilter, Flags)
End Sub

Private Sub Class_Terminate()
   Dim i As Long
   ' Tear down
   For i = 0 To 25
      If m_hDrvNotify(i) Then
         Call UnregisterDeviceNotification(m_hDrvNotify(i))
      End If
   Next i
   Call UnregisterDeviceNotification(m_hDevNotify)
   Call HookClear(m_hWndHidden, Me)
End Sub

' *********************************************
'  Public Methods
' *********************************************
Public Function GetDeviceFlags() As DeviceFlags
   Dim db As DEV_BROADCAST_VOLUME
   Dim i As Long
   ' For use during the Device* events.
   If m_lpDevInfo Then
      ' Retrieve information from cached pointer.
      Call CopyMemory(db, ByVal m_lpDevInfo, Len(db))
      ' Validate structure, then extract value.
      If db.DeviceType = DBT_DEVTYP_VOLUME Then
         GetDeviceFlags = db.Flags
      End If
   End If
End Function

Public Function GetDeviceInterfaceName() As String
   Dim db As DEV_BROADCAST_DEVICEINTERFACE
   ' For use during the Device* events.
   If m_lpDevInfo Then
      ' Retrieve information from cached pointer.
      Call CopyMemory(db, ByVal m_lpDevInfo, Len(db))
      ' Validate structure, then extract string.
      If db.DeviceType = DBT_DEVTYP_DEVICEINTERFACE Then
         GetDeviceInterfaceName = TrimNull(StrConv(db.DeviceName, vbFromUnicode))
      End If
   End If
End Function

Public Function GetDeviceType() As Long
   ' For use during the Device* events.
   If m_lpDevInfo Then
      '   typedef struct _DEV_BROADCAST_HDR {
      '     DWORD dbch_size;
      '     DWORD dbch_devicetype;
      '     DWORD dbch_reserved;
      '   }DEV_BROADCAST_HDR, *PDEV_BROADCAST_HDR;
      GetDeviceType = PointerToDWord(m_lpDevInfo + 4)
   End If
End Function

Public Function GetDeviceVolume() As String
   Dim dbv As DEV_BROADCAST_VOLUME
   Dim dbh As DEV_BROADCAST_HANDLE
   Dim i As Long
   If m_lpDevInfo Then
      Select Case Me.GetDeviceType
         Case DBT_DEVTYP_VOLUME
            ' Retrieve information from cached pointer.
            Call CopyMemory(dbv, ByVal m_lpDevInfo, Len(dbv))
            ' Validate structure, then extract value.
            If dbv.DeviceType = DBT_DEVTYP_VOLUME Then
               For i = 0 To 25
                  If (2 ^ i And dbv.UnitMask) Then
                     GetDeviceVolume = Chr$(65 + i)
                     Exit For
                  End If
               Next i
            End If
         Case DBT_DEVTYP_HANDLE
            ' Retrieve information from cached pointer.
            Call CopyMemory(dbh, ByVal m_lpDevInfo, Len(dbh))
            For i = 0 To 25
               If m_hDrvNotify(i) = dbh.hDevNotify Then
                  GetDeviceVolume = Chr$(65 + i)
                  Exit For
               End If
            Next i
      End Select
   End If
End Function

Public Function SetQueryHook(ByVal Drive As Variant, Optional ByVal SetHook As Boolean = True) As Boolean
   Dim nDrive As Long
   Dim hDrive As Long
   Dim dbh As DEV_BROADCAST_HANDLE
   Const OPEN_ALWAYS = 4
   Const INVALID_HANDLE_VALUE As Long = -1&

   ' Try to make sense of value passed by client.
   If IsNumeric(Drive) Then
      nDrive = Drive
   ElseIf Len(Drive) Then
      nDrive = Asc(UCase$(Left$(Drive, 1))) - 65
   Else
      nDrive = -1
   End If

   If nDrive >= 0 And nDrive <= 25 Then
      ' Clear existing hook, if exists.
      If m_hDrvNotify(nDrive) Then
         Call UnregisterDeviceNotification(m_hDrvNotify(nDrive))
         m_hDrvNotify(nDrive) = 0
      End If

      ' Allow routine to double-duty as a clear hook, too.
      If SetHook Then
         ' Open handle to device, and register callback.
         hDrive = CreateFile("\\.\" & Chr$(65 + nDrive) & ":", 0&, 0&, ByVal 0&, OPEN_ALWAYS, 0&, 0&)
         If hDrive <> INVALID_HANDLE_VALUE Then
            dbh.cbSize = Len(dbh)
            dbh.Handle = hDrive
            dbh.DeviceType = DBT_DEVTYP_HANDLE
            m_hDrvNotify(nDrive) = RegisterDeviceNotification(m_hWndHidden, dbh, DEVICE_NOTIFY_WINDOW_HANDLE)
            SetQueryHook = (m_hDrvNotify(nDrive) <> 0)
            Call CloseHandle(hDrive)
         End If
      Else
         ' Existing hook, if any, should be clear now.
         SetQueryHook = True
      End If
   End If
End Function

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
   Dim EatIt As Boolean
   Dim nRet As Long
   Dim msg As String
   
   ' Special processing for messages we care about.
   Select Case uiMsg
      Case WM_DEVICECHANGE
         Select Case wParam
            Case DBT_DEVICEARRIVAL To DBT_DEVICEREMOVECOMPLETE
               ' Delegate handling for specific events.
               nRet = HandleDeviceChange(wParam, lParam)
               If nRet Then
                  EatIt = True
                  IHookXP_Message = nRet
               End If
            
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
   End Select
   
   ' Pass back to default message handler.
   If EatIt = False Then
      IHookXP_Message = HookDefault(hWnd, uiMsg, wParam, lParam)
   End If

End Function

Private Function HandleDeviceChange(ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim Cancel As Boolean
   Dim DevType As Long
   Dim sDrive As String
   
   ' Determine device type, by looking at pointer+4...
   '   typedef struct _DEV_BROADCAST_HDR {
   '     DWORD dbch_size;
   '     DWORD dbch_devicetype;
   '     DWORD dbch_reserved;
   '   }DEV_BROADCAST_HDR, *PDEV_BROADCAST_HDR;
   DevType = PointerToDWord(lParam + 4)

   ' Cache pointer for duration of call.
   m_lpDevInfo = lParam
   
   Select Case wParam
      ' Provide methods to determine structure contents
      ' which client may call during raised event.
      Case DBT_DEVICEARRIVAL
         ' Hook into removal query chain, then alert client.
         If DevType = DBT_DEVTYP_VOLUME Then
            Call SetQueryHook(Me.GetDeviceVolume, True)
         End If
         RaiseEvent DeviceArrival(DevType)
         
      Case DBT_DEVICEQUERYREMOVE
         ' Allow client to cancel device removal.
         RaiseEvent DeviceQueryRemove(DevType, Cancel)
         If Cancel Then
            HandleDeviceChange = BROADCAST_QUERY_DENY
         End If
         
      Case DBT_DEVICEQUERYREMOVEFAILED
         RaiseEvent DeviceQueryRemoveFailed(DevType)
         
      Case DBT_DEVICEREMOVEPENDING
         ' Alert client, then unhook from query chain.
         RaiseEvent DeviceRemovePending(DevType)
         If DevType = DBT_DEVTYP_VOLUME Then
            Call SetQueryHook(Me.GetDeviceVolume, False)
         End If
         
      Case DBT_DEVICEREMOVECOMPLETE
         RaiseEvent DeviceRemoveComplete(DevType)
   
   End Select
   
   ' Clear cached pointer.
   m_lpDevInfo = 0
End Function

