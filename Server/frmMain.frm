VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Remote USB -Server"
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":1CFA
   ScaleHeight     =   2490
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4080
      Top             =   2400
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3064
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3299
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   6703
      EndProperty
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   0
      Left            =   7440
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Register"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":373D
      cBack           =   -2147483633
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   6
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Exit"
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
      Image           =   "frmMain.frx":3B68
      cBack           =   -2147483633
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "About"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":3DDF
      cBack           =   -2147483633
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Reload USB Devices List"
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Refresh"
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
      Image           =   "frmMain.frx":403C
      cBack           =   -2147483633
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Share USB Device"
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Share"
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
      Image           =   "frmMain.frx":4450
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   5
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Remove Share On USB Device"
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "UnShare"
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
      Image           =   "frmMain.frx":46AD
      cBack           =   -2147483633
   End
   Begin VB.Label lblNav 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VPN Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   0
      Left            =   5040
      MouseIcon       =   "frmMain.frx":48FE
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Make a VPN Network"
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ready."
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
      Height          =   1155
      Index           =   2
      Left            =   4320
      TabIndex        =   9
      Top             =   600
      Width           =   2565
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading USB Devices..."
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
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remote USB -Server"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1740
   End
   Begin VB.Image Image2 
      Height          =   1860
      Left            =   4560
      Picture         =   "frmMain.frx":4A50
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastSelected As String
Private WithEvents sid As CSysInfoDevice
Attribute sid.VB_VarHelpID = -1

Dim WithEvents ShareDeviceApp As ConsoleCapture
Attribute ShareDeviceApp.VB_VarHelpID = -1
Dim WithEvents LoadDevicesApp As ConsoleCapture
Attribute LoadDevicesApp.VB_VarHelpID = -1
Dim bLocked As Boolean      'Locked?
Dim MakeShareOperation As Boolean

Private Sub Form_Initialize()
  
    InitCommonControlsXP
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove 0, 0, 0, 0
End Sub
Private Sub Form_Load()
    
    If Command = "trustnetwork" Then
        Load frmVPN
        frmVPN.ConnectToVPN
        frmVPN.SetNetworkAsHome
        frmVPN.Disconnect
        Unload frmVPN
        Unload Me
        Exit Sub
    End If
    TryDownloadIPFromInternet
   
    
    
    Image1.Picture = Me.Picture
    Image1.Top = 0
    Image1.Left = 0
    Image1.Width = Me.ScaleWidth
    Image1.Height = Me.ScaleHeight
    LastSelected = ""
    SetForm Me, &HFF00FF
    Label1(0).Caption = "Amerizon Remote Programming -Customer"
     '  Me.AutoRedraw = True
      Hide
    If CheckVer(Label1(0), True) = False Then
        MsgBox "Your trail period is expired." + vbNewLine + "Please contact with:Ashrafnet4u@hotmail.com" + vbNewLine + "Desktop Team", vbCritical
        End
    Else
    
        Show
        Set LoadDevicesApp = New ConsoleCapture
        Set ShareDeviceApp = New ConsoleCapture
        loaddevices
        Timer1.Enabled = True
        Set sid = New CSysInfoDevice
        Load frmVPN
        frmVPN.Visible = False
        frmVPN.ConnectToVPN
    End If
    'tmrWindow.Enabled = True
    
   
End Sub

Sub TryDownloadIPFromInternet()
Exit Sub
Dim sFileText As String
    Dim iFileNo As Integer
  iFileNo = FreeFile
  Dim strTmpFile As String
  Dim IPAddress As String
  strTmpFile = GetTempFolder() & "vpn.dat"
  On Error GoTo er:
 ' Try to get IpAddress from Internet
    If DownloadFile("http://www.amerizonwireless.com/vpn.dat", strTmpFile) Then
        Open strTmpFile For Input As #iFileNo
          Do While Not EOF(iFileNo)
            Input #iFileNo, sFileText
             If sFileText <> "" & Len(sFileText) > 6 Then
                IPAddress = Trim(sFileText)
            End If
          Loop
        Close #iFileNo
    End If
    
    If IPAddress <> "" Then
        SaveString HKEY_CURRENT_USER, "RemoteUSB", "IPAddress", IPAddress
    End If
 
  
er:
End Sub
Sub ShareDevice(MakeShare As Boolean)
On Error GoTo er:
    
   
    
    Screen.MousePointer = vbHourglass
    MakeShareOperation = MakeShare
    
    Dim curCmd As String
    Dim strShare As String
    strShare = "share"
    If MakeShare = False Then strShare = "unshare"
    curCmd = GetShortName(App.Path & "\usbsrvcmd.dat") & " " & strShare & " " & Chr(34) & ListView1.SelectedItem.Tag & Chr(34)
    
    ShareDeviceApp.RunProcess "cmd /k " & curCmd     'Launch the console
    Screen.MousePointer = vbDefault
    Exit Sub
er:

    Label1(2).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(2).ForeColor = vbRed
    Screen.MousePointer = vbDefault
End Sub

Sub loaddevices()
    On Error GoTo er:
    
    lvButtons_H1(3).Enabled = False
    ListView1.ListItems.Clear
    Label1(1).Caption = "Loading USB Devices..."
    DoEvents
    Screen.MousePointer = vbHourglass
    DoEvents
        
    Dim curCmd As String
    curCmd = GetShortName(App.Path & "\usbsrvcmd.dat") & " list"
    
    LoadDevicesApp.RunProcess "cmd /k " & curCmd     'Launch the console
    
    If ListView1.ListItems.Count > 0 Then
        selectlastone
        lvButtons_H1(3).Enabled = True
        Label1(1).Caption = "Available USB Devices [" & ListView1.ListItems.Count & " Device/s]"
        
    Else
        Label1(2).ForeColor = vbRed: Label1(2).FontBold = True
        Label1(2) = "No USB devices attached to your PC."
        Label1(1).Caption = "Available USB Devices [0 Device]"
        lvButtons_H1(3).Enabled = True
        lvButtons_H1(4).Enabled = False
        lvButtons_H1(5).Enabled = False
'        ListView1.Enabled = False: lvButtons_H1(0).Enabled = False: lvButtons_H1(3).Enabled = False: lvButtons_H1(4).Enabled = False: lvButtons_H1(5).Enabled = False
'        Label1(1).Caption = ""
    End If
     
    
    If ListView1.ListItems.Count > 0 Then
        ListView1_ItemClick ListView1.SelectedItem
    End If
    CheckStateDevices
    Screen.MousePointer = vbDefault
Exit Sub
er:
    Label1(2).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(2).ForeColor = vbRed
    Screen.MousePointer = vbDefault
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub

Private Sub LoadDevicesApp_ReadProcess(readData As String)
    On Error GoTo er:
    'MsgBox readData
  '  readData = Trim(readData)
'    Clipboard.Clear: Clipboard.SetText readData
'    readData = Replace(readData, vbNewLine, ";")
    If Left(readData, 1) = ";" Then
        readData = Mid(readData, 2)
    End If
    Dim yy() As String
    yy = Split(readData, vbNewLine)
    Dim strDev As String
    Dim i As Integer
    For i = 0 To UBound(yy)
        strDev = yy(i)
        strDev = Trim(strDev)
        If strDev = "" Then GoTo nxt:
        Dim Name As String
        If Left(strDev, 3) = "USB" Then
            
            Name = Mid(strDev, InStr(1, strDev, "- ") + 5)
            If InStr(1, Name, ";") Then Name = Mid(Name, 1, InStr(1, Name, ";") - 1)
            Name = Trim(Name)
            If Name <> "" Then
                Dim itmX As ListItem
                Set itmX = ListView1.ListItems.Add(, , Name, 1, 1)
                itmX.Tag = Trim(Mid(strDev, 1, InStr(1, strDev, "- ") - 1))
            End If
            
        End If
nxt:
    Next i
        
    Exit Sub
er:
    Screen.MousePointer = vbDefault
    Label1(2).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(2).ForeColor = vbRed
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown Button, Shift, x, y
End Sub

Private Sub lblNav_Click(Index As Integer)
    frmVPN.Show , Me
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo er:
    LastSelected = Item.Text

    Dim State As String
    State = GetSetting(App.CompanyName, "Devices", ListView1.SelectedItem.Tag, 0)
    lvButtons_H1(4).Enabled = False
    lvButtons_H1(5).Enabled = False
    Label1(2).Caption = Item.Text
    Label1(2).FontBold = False
    If State = 0 Then
        SetStatus Label1(2).Caption & vbNewLine & "Not Shared yet."
        Label1(2).ForeColor = vbBlack
        lvButtons_H1(4).Enabled = True
        Item.SmallIcon = 1
    ElseIf State = 1 Then
        SetStatus Label1(2).Caption & " [Shared]" & vbNewLine & ""
        Label1(2).ForeColor = vbBlue
        lvButtons_H1(5).Enabled = True
        Item.SmallIcon = 2
    ElseIf State = 2 Then
        SetStatus Label1(2).Caption & " [Connected]"
        Label1(2).ForeColor = &H8000&
        Label1(2).FontBold = True
        lvButtons_H1(5).Enabled = True
        Item.SmallIcon = 3
    End If
    Exit Sub
er:
    Label1(2).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(2).ForeColor = vbRed
End Sub

Sub SetStatus(strStatus As String)
    If Label1(2).Caption = strStatus Then Exit Sub
    Label1(2).Caption = strStatus
End Sub
Sub CheckStateDevices()
    On Error GoTo er:
    Dim i As Integer
    
Dim State As String
For i = 1 To ListView1.ListItems.Count
    State = 0
    State = GetSetting(App.CompanyName, "Devices", ListView1.ListItems(i).Tag, 0)

    If State = 0 Then
        ListView1.ListItems(i).SmallIcon = 1
    ElseIf State = 1 Then
        ListView1.ListItems(i).SmallIcon = 2
    ElseIf State = 2 Then
        ListView1.ListItems(i).SmallIcon = 3
    End If
Next i
    
    
    If ListView1.ListItems.Count > 0 Then
        ListView1_ItemClick ListView1.SelectedItem
    End If
    Exit Sub
er:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)
    Dim IsSuccess As Boolean
    Select Case Index
        Case 0 'register
            frmReg.Show , Me
        Case 1 'Close App
            Unload frmVPN
            Unload Me
            Exit Sub
        Case 2 ' About form
            frmAbout.Show , Me
        Case 3 ' Refresh
            loaddevices
        Case 4 ' Share Device
                ShareDevice True
        Case 5 ' Unshare Device
           ShareDevice False
        Case 6 ' Get IP
            Dim IpAddrs
            IpAddrs = GetIpAddrTable
            Debug.Print "Nr of IP addresses: " & UBound(IpAddrs) - LBound(IpAddrs) + 1
            Dim i As Integer
            For i = LBound(IpAddrs) To UBound(IpAddrs)
                Debug.Print IpAddrs(i)
            Next

    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblNav(0).FontUnderline = False
    lblNav(0).ForeColor = vbBlue
End Sub

Private Sub lblNav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not lblNav(Index).FontUnderline Then
        lblNav(Index).FontUnderline = True
        lblNav(Index).ForeColor = vbRed
    End If
End Sub
Sub selectlastone()
On Error GoTo er:
    Dim i As Integer
    For i = 1 To ListView1.ListItems.Count
        If LastSelected = ListView1.ListItems(i).Text Then
            ListView1.ListItems(i).Selected = True
            Exit Sub
        End If
    Next i
er:
End Sub




Private Sub ShareDeviceApp_ReadProcess(readData As String)
    On Error GoTo er:
    readData = Trim(readData)
    readData = Replace(readData, vbNewLine, ";")
    ShareDeviceApp.Cancel
    
    If InStr(1, readData, "FAILED") Then
        If InStr(1, readData, "UsbEngineServer (2192)") Then
            readData = "OK"
            GoTo TryOK:
        End If
        
        
        Label1(2).Caption = "Error: " & vbNewLine & "" & "USB Device Is not Pluged in."
        Label1(2).ForeColor = vbRed
        lvButtons_H1(4).Enabled = True
        Exit Sub
    End If
TryOK:
    If InStr(1, readData, "OK") Then
        lvButtons_H1(4).Enabled = Not MakeShareOperation
        lvButtons_H1(5).Enabled = MakeShareOperation
        SaveSetting App.CompanyName, "Devices", ListView1.SelectedItem.Tag, IIf(MakeShareOperation, 1, 0)
    Else
        lvButtons_H1(5).Enabled = Not MakeShareOperation
        lvButtons_H1(4).Enabled = MakeShareOperation
    End If

        
    Exit Sub
er:
    Screen.MousePointer = vbDefault
    Label1(2).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(2).ForeColor = vbRed
End Sub

Private Sub sid_DeviceArrival(ByVal DevType As DeviceTypes)
    lvButtons_H1_Click 3
End Sub

Private Sub sid_DeviceRemoveComplete(ByVal DevType As DeviceTypes)
    lvButtons_H1_Click 3
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    CheckStateDevices
    Timer1.Enabled = True
End Sub

