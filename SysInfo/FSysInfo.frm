VERSION 5.00
Begin VB.Form FSysInfo 
   Caption         =   "SysInfo Demo"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEvents 
      Height          =   975
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "FSysInfo.frx":0000
      Top             =   360
      Width           =   1755
   End
End
Attribute VB_Name = "FSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private WithEvents sig As CSysInfoGeneral
Attribute sig.VB_VarHelpID = -1
Private WithEvents sid As CSysInfoDevice
Attribute sid.VB_VarHelpID = -1
Private WithEvents sip As CSysInfoPower
Attribute sip.VB_VarHelpID = -1

Private Sub Form_Load()
   ' Attach to global events
   Set sig = g_SysInfoGeneral
   Set sid = g_SysInfoDevice
   Set sip = g_SysInfoPower
   ' Configure controls
   txtEvents.Text = ""
End Sub

Private Sub Form_Resize()
   txtEvents.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub sid_DeviceArrival(ByVal DevType As DeviceTypes)
   ShowEvent BuildDeviceMessage("DeviceArrival")
End Sub

Private Sub sid_DeviceQueryRemove(ByVal DevType As DeviceTypes, Cancel As Boolean)
   ShowEvent BuildDeviceMessage("DeviceQueryRemove")
   'Cancel = True
End Sub

Private Sub sid_DeviceQueryRemoveFailed(ByVal DevType As DeviceTypes)
   ShowEvent BuildDeviceMessage("DeviceQueryRemoveFailed")
End Sub

Private Sub sid_DeviceRemoveComplete(ByVal DevType As DeviceTypes)
   ShowEvent BuildDeviceMessage("DeviceRemoveComplete")
End Sub

Private Sub sid_DeviceRemovePending(ByVal DevType As DeviceTypes)
   ShowEvent BuildDeviceMessage("DeviceRemovePending")
End Sub

Private Sub sig_ActivateApp(ByVal Active As Boolean)
   ShowEvent "ActivateApp: " & IIf(Active, "Active", "Inactive")
End Sub

Private Sub sig_DisplayChange(ByVal BitsPerPixel As Long, ByVal PixelsX As Long, ByVal PixelsY As Long)
   ShowEvent "DisplayChange: " & CStr(PixelsX) & " by " & CStr(PixelsY) & " pixels, " _
                               & CStr(BitsPerPixel) & "bpp"
End Sub

Private Sub sig_EndSession(ByVal EndingInitiated As Boolean, ByVal Flag As Long)
   ShowEvent "EndSession: " & CStr(EndingInitiated) & ", Flag=&h" & Hex$(Flag)
End Sub

Private Sub sig_FontChange()
   ShowEvent "FontChange: " & Now
End Sub

Private Sub sig_QueryEndSession(ByVal Flag As EndSessionFlags, Cancel As Boolean)
   ShowEvent "QueryEndSession: Flag=&h" & Hex$(Flag)
   'Cancel = True
End Sub

Private Sub sig_SettingChange(ByVal Setting As String, ByVal Flag As Long)
   ShowEvent "SettingChange: " & Setting & " (flag: &h" & Hex$(Flag) & ")"
End Sub

Private Sub sig_SysColorChange()
   ShowEvent "SysColorChange: " & Now
End Sub

Private Sub sig_ThemeChanged()
   ShowEvent "ThemeChanged: " & Now
End Sub

Private Sub sig_TimeChanged()
   ShowEvent "TimeChanged: " & Now
End Sub

Private Sub sip_PowerBatteryLow()
   ShowEvent BuildPowerMessage("PowerBatteryLow")
End Sub

Private Sub sip_PowerResume()
   ShowEvent BuildPowerMessage("PowerResume")
End Sub

Private Sub sip_PowerResumeAutomatic()
   ShowEvent BuildPowerMessage("PowerResumeAutomatic")
End Sub

Private Sub sip_PowerResumeCritical()
   ShowEvent BuildPowerMessage("PowerResumeCritical")
End Sub

Private Sub sip_PowerSettingChange(ByVal lpSetting As Long)
   ShowEvent BuildPowerMessage("PowerSettingChange")
End Sub

Private Sub sip_PowerStatusChange()
   ShowEvent BuildPowerMessage("PowerStatusChange")
End Sub

Private Sub sip_PowerSuspend()
   ShowEvent BuildPowerMessage("PowerSuspend")
End Sub

Private Sub sip_PowerSuspendQuery(Cancel As Boolean)
   ShowEvent BuildPowerMessage("PowerSuspendQuery")
End Sub

Private Sub sip_PowerSuspendQueryFailed()
   ShowEvent BuildPowerMessage("PowerSuspendQueryFailed")
End Sub

Private Sub ShowEvent(ByVal EventText As String)
   With txtEvents
      EventText = Format$(Now, "hh:mm:ss") & " - " & EventText
      .SelStart = Len(.Text)
      .SelText = EventText & vbCrLf
   End With
   #If Debugging Then
      Print #hLog, EventText
   #End If
End Sub

Private Function BuildDeviceMessage(ByVal EventName As String) As String
   Dim msg As String
   msg = EventName & ": " & sid.GetDeviceType & " - "
   Select Case sid.GetDeviceType
      Case DeviceTypeVolume
         msg = msg & "vol: " & sid.GetDeviceVolume & ", flags: " & Hex$(sid.GetDeviceFlags)
      Case DeviceTypeHandle
         msg = msg & "vol: " & sid.GetDeviceVolume
      Case DeviceTypeInterface
         msg = msg & sid.GetDeviceInterfaceName
   End Select
   BuildDeviceMessage = msg
End Function

Private Function BuildPowerMessage(ByVal EventName As String) As String
   Dim msg As String
   msg = EventName & ":"
   
   Select Case sip.ACLineStatus
      Case 0
         msg = msg & " AC=no"
      Case 1
         msg = msg & " AC=yes"
      Case 255
         msg = msg & " AC=(unknown)"
   End Select
   
   If sip.BatteryLifeTime >= 0 Then
      msg = msg & " Life=" & Format$(sip.BatteryLifeTime \ 60) & "min."
   End If
   
   If sip.BatteryLifePercent < 255 Then
      msg = msg & " Percent=" & CStr(sip.BatteryLifePercent) & "%"
   Else
      msg = msg & " Percent=(unknown)"
   End If
   
   msg = msg & " Flags=" & CStr(sip.BatteryFlags)
   BuildPowerMessage = msg
End Function

Private Function FormatSeconds(ByVal Seconds As Long) As String
   FormatSeconds = Format$(Seconds \ 60, "0") & ":" & Format$(Seconds Mod 60, "00")
End Function
