VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSysInfoPower"
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
Private Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type SYSTEM_POWER_STATUS
   ACLineStatus As Byte
   BatteryFlag As Byte
   BatteryLifePercent As Byte
   Reserved1 As Byte
   BatteryLifeTime As Long
   BatteryFullLifeTime As Long
End Type

' Monitored messages.
Private Const WM_POWERBROADCAST               As Long = &H218

Private Const PBT_APMQUERYSUSPEND             As Long = &H0
Private Const PBT_APMQUERYSTANDBY             As Long = &H1
Private Const PBT_APMQUERYSUSPENDFAILED       As Long = &H2
Private Const PBT_APMQUERYSTANDBYFAILED       As Long = &H3
Private Const PBT_APMSUSPEND                  As Long = &H4
Private Const PBT_APMSTANDBY                  As Long = &H5
Private Const PBT_APMRESUMECRITICAL           As Long = &H6
Private Const PBT_APMRESUMESUSPEND            As Long = &H7
Private Const PBT_APMRESUMESTANDBY            As Long = &H8

Private Const PBTF_APMRESUMEFROMFAILURE       As Long = &H1
Private Const PBT_APMBATTERYLOW               As Long = &H9
Private Const PBT_APMPOWERSTATUSCHANGE        As Long = &HA
Private Const PBT_APMOEMEVENT                 As Long = &HB
Private Const PBT_APMRESUMEAUTOMATIC          As Long = &H12

Private Const PBT_POWERSETTINGCHANGE          As Long = &H8013&

' Return this value to deny a query.
Private Const BROADCAST_QUERY_DENY            As Long = &H424D5144

' Subclassing interface
Implements IHookXP

' System notification events.
Public Event PowerBatteryLow()
Public Event PowerResume()
Public Event PowerResumeAutomatic()
Public Event PowerResumeCritical()
Public Event PowerSettingChange(ByVal lpSetting As Long)
Public Event PowerStatusChange()
Public Event PowerSuspend()
Public Event PowerSuspendQuery(Cancel As Boolean)
Public Event PowerSuspendQueryFailed()

Public Enum PowerACStatus
   ACOffline = 0
   ACOnline = 1
   ACUnknown = 2
End Enum

Public Enum PowerBatteryFlags
   BatteryCapacityHigh = 1      ' High—the battery capacity is at more than 66 percent
   BatteryCapacityLow = 2       ' Low—the battery capacity is at less than 33 percent
   BatteryCapacityCritical = 4  ' Critical—the battery capacity is at less than five percent
   BatteryCapacityCharging = 8  ' Charging
   BatteryNone = 128            ' No system battery
   BatteryUnknown = 255         ' Unknown status—unable to read the battery flag information
End Enum

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
Public Function ACLineStatus() As PowerACStatus
   Dim sps As SYSTEM_POWER_STATUS
   If GetSystemPowerStatus(sps) Then
      ACLineStatus = sps.ACLineStatus
   End If
End Function

Public Function BatteryFlags() As PowerBatteryFlags
   Dim sps As SYSTEM_POWER_STATUS
   If GetSystemPowerStatus(sps) Then
      BatteryFlags = sps.BatteryFlag
   End If
End Function

Public Function BatteryLifePercent() As Long
   Dim sps As SYSTEM_POWER_STATUS
   ' The percentage of full battery charge remaining. This member can be a
   ' value in the range 0 to 100, or 255 if status is unknown.
   If GetSystemPowerStatus(sps) Then
      BatteryLifePercent = sps.BatteryLifePercent
   End If
End Function

Public Function BatteryLifeTime() As Long
   Dim sps As SYSTEM_POWER_STATUS
   ' The number of seconds of battery life remaining, or –1 if remaining
   ' seconds are unknown.
   If GetSystemPowerStatus(sps) Then
      BatteryLifeTime = sps.BatteryLifeTime
   End If
End Function

Public Function BatteryFullLifeTime() As Long
   Dim sps As SYSTEM_POWER_STATUS
   ' The number of seconds of battery life when at full charge, or –1 if
   ' full battery lifetime is unknown.
   If GetSystemPowerStatus(sps) Then
      BatteryFullLifeTime = sps.BatteryFullLifeTime
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
   Dim Cancel As Boolean
   Dim EatIt As Boolean
   Dim msg As String
   
   ' Special processing for messages we care about.
   Select Case uiMsg
      Case WM_POWERBROADCAST
         Select Case wParam
            Case PBT_APMBATTERYLOW
               ' Notifies applications that the battery power is low.
               RaiseEvent PowerBatteryLow
               
            Case PBT_APMRESUMESUSPEND
               ' Notifies applications that the system has resumed
               ' operation after being suspended.
               RaiseEvent PowerResume
               
            Case PBT_APMRESUMEAUTOMATIC
               ' Notifies applications that the computer has woken up
               ' automatically to handle an event. An application will
               ' not generally respond unless it is handling the event,
               ' because the user is not present.
               RaiseEvent PowerResumeAutomatic
               
            Case PBT_APMRESUMECRITICAL
               ' Notifies applications that the system has resumed operation.
               ' This event can indicate that some or all applications did
               ' not receive a PBT_APMSUSPEND event. For example, this event
               ' can be broadcast after a critical suspension caused by a
               ' failing battery.
               RaiseEvent PowerResumeCritical
               
            Case PBT_POWERSETTINGCHANGE
               ' Power setting change, lParam points to structure.
               RaiseEvent PowerSettingChange(lParam)
               
            Case PBT_APMPOWERSTATUSCHANGE
               ' Notifies applications of a change in the power status of the
               ' computer, such as a switch from battery power to A/C. The
               ' system also broadcasts this event when remaining battery power
               ' slips below the threshold specified by the user or if the
               ' battery power changes by a specified percentage.
               RaiseEvent PowerStatusChange
               
            Case PBT_APMSUSPEND
               ' Notifies applications that the computer is about to enter a
               ' suspended state. This event is typically broadcast when all
               ' applications and installable drivers have returned TRUE to a
               ' previous PBT_APMQUERYSUSPEND event.
               RaiseEvent PowerSuspend
               
            Case PBT_APMQUERYSUSPEND
               ' Requests permission to suspend the computer. An
               ' application that grants permission should carry out
               ' preparations for the suspension before returning.
               RaiseEvent PowerSuspendQuery(Cancel)
               If Cancel Then
                  EatIt = True
                  IHookXP_Message = BROADCAST_QUERY_DENY
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


