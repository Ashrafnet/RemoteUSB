VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmVPN 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "VPN Connection"
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   Picture         =   "frmVPN.frx":0000
   ScaleHeight     =   1845
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmVPN.frx":117E
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   33002
   End
   Begin VB.ComboBox cboDevices 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3720
      Top             =   1440
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   0
      ToolTipText     =   "Close"
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   255
      cFHover         =   128
      LockHover       =   2
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVPN.frx":129C
      cBack           =   -2147483633
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Close a VPN Connection"
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Disconnect"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   8388608
      LockHover       =   2
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVPN.frx":1513
      cBack           =   -2147483633
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Make a VPN Connection"
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Connect"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   32768
      cFHover         =   16384
      LockHover       =   2
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmVPN.frx":178A
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ready."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   4785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use this tool to make a VPN Network between you and AmerizonWireless Co."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   4755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remote USB -VPN Connection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmVPN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents conApp As ConsoleCapture
Attribute conApp.VB_VarHelpID = -1
Dim WithEvents IPVPNApp As ConsoleCapture
Attribute IPVPNApp.VB_VarHelpID = -1
Dim bLocked As Boolean      'Locked?


Private VPNname As String
Private DeviceName As String
Private Connected As Boolean

Public VPNIP As String
Dim IPSent As Boolean

Const ERROR_NO_MORE_ITEMS = 259&
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_LOCAL_MACHINE = &H80000002

Function GetAdapterIP(AdpterName As String)
    Dim objWMIService, colItems, objItem, macadd, arrIPAddresses, yy, strAddress
    Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery _
        ("Select * From Win32_NetworkAdapter Where NetConnectionID like '%" & AdpterName & "%'")
    
    For Each objItem In colItems
        If Not IsNull(objItem.MACAddress) Then
            macadd = objItem.MACAddress
        End If
    Next
'    Clipboard.Clear: Clipboard.SetText AdpterName
    Set colItems = objWMIService.ExecQuery _
        ("Select * From Win32_NetworkAdapterConfiguration where IPEnabled=True")
    
    For Each objItem In colItems
        arrIPAddresses = objItem.MACAddress
        If arrIPAddresses = macadd Then
            yy = objItem.IPAddress
            If IsNull(yy) Then GoTo nxt:
            For Each strAddress In yy
                If IsNumeric(Mid(strAddress, 1, 2)) Then
                    GetAdapterIP = strAddress
                    Exit Function
                End If
            Next
        End If
        
nxt:
    Next
End Function

Private Sub ListDevices()
    Dim i As Integer
    cboDevices.Clear
    GetDevices cboDevices
    If cboDevices.ListCount > 0 Then
        cboDevices.ListIndex = 0
        For i = 0 To cboDevices.ListCount - 1
            If InStr(1, LCase((cboDevices.List(i))), "ikev2") Then
                DeviceName = cboDevices.List(i)
                Exit Sub
            End If
        Next i
        
        For i = 0 To cboDevices.ListCount - 1
            If InStr(1, LCase((cboDevices.List(i))), "pptp") Then
                DeviceName = cboDevices.List(i)
                Exit Sub
            End If
        Next i
        
    Else
        lvButtons_H1(2).Enabled = False
    End If
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown 0, 0, 0, 0
End Sub


Private Sub Form_Load()
    VPNname = "AmerizonWireless"
    If Command = "trustnetwork" Then
        ListDevices
        Exit Sub
    End If
    Image1.Picture = Me.Picture
    Image1.Top = 0
    Image1.Left = 0
    Image1.Width = Me.ScaleWidth
    Image1.Height = Me.ScaleHeight
    SetForm Me, &HFF00FF
    Hide
    ListDevices
'    Show
    Set conApp = New ConsoleCapture
    Set IPVPNApp = New ConsoleCapture
    
    'VPNname = "AmerizonWireless"
    Label1(0).Caption = "Amerizon Remote Programming -VPN Connection"
    checkstatus

End Sub

Sub SetNetworkAsHome()
On Error Resume Next
    Dim filelocation As String
    Dim strData As String
    strData = Text1.Text
    strData = Replace(strData, "%VPNName%", VPNname)
    filelocation = App.Path & "\net.ps1"
    'On Error Resume Next
    '
    Open filelocation For Output As #1
        Print #1, strData
    Close #1
    filelocation = GetShortName(filelocation)
    Dim curCmd As String
    curCmd = "" & "powershell set-executionpolicy remotesigned"
    ShellandWait curCmd
    curCmd = "" & "powershell " & Chr(34) & filelocation & Chr(34)
    'Clipboard.Clear: Clipboard.SetText curCmd
    ShellandWait curCmd
    Kill filelocation
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'clear reference to class
    Set conApp = Nothing                        'Unload conApp
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Public Sub CreateNewEntry()
    'make the DUN called DUN_NAME
    'Will work if the selected device in the combo box is a modem
    'The Connection Created is a "dummy" one
    'If you want a real one, change the below parameters
    'to the what you need, or add text boxes to a form
    'so the user can enter them him/herself
    If DeviceName = "" Then Exit Sub

    Dim typVBRasEntry As VBRasEntry
    typVBRasEntry.AreaCode = ""
    typVBRasEntry.AutodialFunc = 0
    typVBRasEntry.CountryCode = "1"
    typVBRasEntry.CountryID = "1"
    typVBRasEntry.DeviceName = DeviceName
    typVBRasEntry.DeviceType = "Modem"
    typVBRasEntry.fNetProtocols = RASNP_Ip
    typVBRasEntry.FramingProtocol = RASFP_Ppp
    typVBRasEntry.options = RASEO_SwCompression + RASEO_IpHeaderCompression _
    + RASEO_SpecificNameServers



'    typVBRasEntry.ipAddrDns.a = "206"
'    typVBRasEntry.ipAddrDns.b = "211"
'    typVBRasEntry.ipAddrDns.c = "214"
'    typVBRasEntry.ipAddrDns.d = "206"
'    typVBRasEntry.ipAddrDnsAlt.a = "212"
'    typVBRasEntry.ipAddrDnsAlt.b = "200"
'    typVBRasEntry.ipAddrDnsAlt.c = "200"
'    typVBRasEntry.ipAddrDnsAlt.d = "200"
'    typVBRasEntry.ipAddrWins.a = "0"
'    typVBRasEntry.ipAddrWins.b = "0"
'    typVBRasEntry.ipAddrWins.c = "0"
'    typVBRasEntry.ipAddrWins.d = "0"
'    typVBRasEntry.ipAddrWinsAlt.a = "0"
'    typVBRasEntry.ipAddrWinsAlt.b = "0"
'    typVBRasEntry.ipAddrWinsAlt.c = "0"
'    typVBRasEntry.ipAddrWinsAlt.d = "0"
    ' ??? ??????
    typVBRasEntry.LocalPhoneNumber = GetIPAddress
   ' typVBRasEntry.LocalPhoneNumber = "akram-pc"

    Dim rtn As Long
    rtn = VBRasSetEntryProperties(VPNname, typVBRasEntry)
    If rtn <> 0 Then
        MsgBox VBRASErrorHandler(rtn)
    Else
    

    End If
End Sub

Function GetIPAddress() As String
    Dim x As String
    GetIPAddress = "remotecp200.amerizon.com"
    x = GetString(HKEY_CURRENT_USER, "RemoteUSB", "IPAddress")
    If x <> "" & Len(x) > 6 Then
        GetIPAddress = Trim(x)
        Exit Function
    End If
    
    Dim sFileText As String
    Dim iFileNo As Integer
  iFileNo = FreeFile
  On Error GoTo er:
  Open App.Path & "\vpn.dat" For Input As #iFileNo
       
  Do While Not EOF(iFileNo)
    Input #iFileNo, sFileText
     If sFileText <> "" & Len(sFileText) > 6 Then
        GetIPAddress = Trim(sFileText)
        Exit Function
    End If
  Loop
  Close #iFileNo
  
er:
  
  
End Function


Private Sub IPVPNApp_ReadProcess(readData As String)
    On Error GoTo er:
    Dim y As Integer
    Dim x As Integer
    y = InStr(1, LCase(readData), LCase(VPNname))
    If y > 0 Then
        y = InStr(y + 1, readData, "IPv4 Address")
        y = InStr(y + 1, readData, ":")
        x = InStr(y + 1, readData, Chr(10))
        VPNIP = Trim(Mid(readData, y, x - y))
        VPNIP = Trim(Replace(VPNIP, ":", ""))
        If InStr(1, VPNIP, "(") Then
            VPNIP = Mid(VPNIP, 1, InStr(1, VPNIP, "(") - 1)
            If VPNIP <> "" Then IPVPNApp.Cancel
        End If
       ' MsgBox readData
    End If
  '  MsgBox readData
er:
    Screen.MousePointer = vbDefault

End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown 0, 0, 0, 0
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)
    Select Case Index
        Case 0
            Hide
        Case 1
            Disconnect
        Case 2 ' Create VPN then Connect
            ConnectToVPN
    End Select
    
End Sub

Public Sub ConnectToVPN()
'    Exit Sub
    Disconnect
    CreateNewEntry
    Connect
End Sub

Private Sub conApp_ReadProcess(readData As String)
    On Error GoTo er:
    'Cancel console application
    'conApp.Cancel
        
    If InStr(1, LCase(readData), LCase(VPNname)) > 0 Then
        Connected = True
        conApp.Cancel
    Else
        Connected = False
    End If
        
    lvButtons_H1(1).Enabled = Connected
    lvButtons_H1(2).Enabled = Not Connected
    If Connected Then
        Label1(2).ForeColor = &H8000&
        Label1(2).Caption = "[Connected]"
    Else
        Label1(2).ForeColor = vbRed
        Label1(2).Caption = "Not Connected Yet."
    End If
    
er:
    Screen.MousePointer = vbDefault
End Sub
Sub Disconnect()
    If bLocked = True Then Exit Sub
    lvButtons_H1(2).Enabled = False
    lvButtons_H1(1).Enabled = False
    'Screen.MousePointer = vbHourglass
    Dim curCmd As String
    bLocked = True                                  'Lock the output window

    curCmd = "" & "rasphone -h  " & Chr(34) & VPNname & Chr(34)
    ShellandWait curCmd
'    curCmd = "" & "rasphone -r  " & Chr(34) & VPNname & Chr(34)
'    ShellandWait curCmd
    'conApp.RunProcess "cmd /k " & curCmd     'Launch the console
    bLocked = False                                 'Unlock
    checkstatus
    Screen.MousePointer = vbDefault
End Sub
Sub Connect()
On Error GoTo er:
    If bLocked = True Then Exit Sub
    lvButtons_H1(2).Enabled = False
    lvButtons_H1(1).Enabled = False
    'Screen.MousePointer = vbHourglass
    Dim curCmd As String
    Dim xx As Boolean
    bLocked = True                                  'Lock the output window
    
    curCmd = "" & "rasdial " & Chr(34) & VPNname & Chr(34) & " vpnuser Amerizon1"
    xx = ShellandWait(curCmd)
    checkstatus
    
er:
    bLocked = False                                 'Unlock
    Screen.MousePointer = vbDefault
End Sub



Sub checkstatus()
On Error GoTo er:
    Dim curCmd As String
    
    curCmd = "rasdial"
    conApp.RunProcess "cmd /k " & curCmd     'Launch the console
    
    If Not Connected Then Exit Sub
    
    'VPNIP = GetAdapterIP(VPNname)
    If VPNIP = "" Then Exit Sub
    If IPSent Then Exit Sub
    If Winsock1.State Then Winsock1.Close
    Winsock1.LocalPort = 0
    Winsock1.Connect "71.28.118.6", 33001
    Exit Sub
er:
   ' Resume
  '  MsgBox Err.Description
   ' Clipboard.Clear: Clipboard.SetText Err.Description
   ' Connected = False
End Sub


Private Sub Timer1_Timer()
    checkstatus
    Dim curCmd As String
    If Connected Then
        curCmd = "ipconfig /all"
        IPVPNApp.RunProcess "cmd /k " & curCmd     'Launch the console
    End If
End Sub

Private Sub Winsock1_Connect()
    If Winsock1.State = sckConnected And IPSent = False Then
        If VPNIP = "" Then Exit Sub
        Winsock1.SendData VPNIP
        IPSent = True
       
    End If
    
End Sub

