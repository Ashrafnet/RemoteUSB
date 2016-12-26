VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Remote USB -Client"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1CFA
   ScaleHeight     =   3600
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   4800
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5880
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   33001
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   333
      Left            =   1080
      TabIndex        =   0
      Text            =   "Localhost"
      Top             =   960
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6480
      Top             =   2280
   End
   Begin RemoteUSBClient.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
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
      Image           =   "frmMain.frx":3202
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
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
   Begin RemoteUSBClient.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   6
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
      cFHover         =   128
      LockHover       =   2
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmMain.frx":3479
      cBack           =   -2147483633
   End
   Begin RemoteUSBClient.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   5
      Top             =   120
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
      cFHover         =   8388608
      LockHover       =   2
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":36D6
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B01
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F93
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RemoteUSBClient.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
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
      Image           =   "frmMain.frx":442B
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin RemoteUSBClient.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   5
      Left            =   1680
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
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
      cFore           =   255
      cFHover         =   128
      LockHover       =   2
      cGradient       =   0
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":487E
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin RemoteUSBClient.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   10
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Find"
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
      Image           =   "frmMain.frx":4AF5
      cBack           =   -2147483633
   End
   Begin RemoteUSBClient.lvButtons_H lvButtons_H1 
      Height          =   375
      Index           =   6
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   "Remove Selected Usb Device."
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Remove"
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
      Image           =   "frmMain.frx":4D51
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label lblNav 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Refresh]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   2
      Left            =   3720
      MouseIcon       =   "frmMain.frx":4FC8
      MousePointer    =   99  'Custom
      TabIndex        =   14
      ToolTipText     =   "Refresh Usb Devices."
      Top             =   615
      Width           =   690
   End
   Begin VB.Label lblNav 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remotely"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   1200
      MouseIcon       =   "frmMain.frx":511A
      MousePointer    =   99  'Custom
      TabIndex        =   13
      ToolTipText     =   "Find remote usb devices and connect"
      Top             =   480
      Width           =   915
   End
   Begin VB.Label lblNav 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recently"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   240
      MouseIcon       =   "frmMain.frx":526C
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "Recently Usb Devices"
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remote IP"
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
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1000
      Width           =   750
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
      Height          =   1395
      Index           =   1
      Left            =   4560
      TabIndex        =   9
      Top             =   1440
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amerizon Remote Programming -Client"
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
      TabIndex        =   8
      Top             =   120
      Width           =   3330
   End
   Begin VB.Image Image2 
      Height          =   2340
      Left            =   4560
      Picture         =   "frmMain.frx":53BE
      Top             =   840
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1755
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents LoadDevicesApp As ConsoleCapture
Attribute LoadDevicesApp.VB_VarHelpID = -1
Dim WithEvents AddedDevicesApp As ConsoleCapture
Attribute AddedDevicesApp.VB_VarHelpID = -1
Dim WithEvents ConnectDevicesApp As ConsoleCapture
Attribute ConnectDevicesApp.VB_VarHelpID = -1

Dim MakeConnectOperation As Boolean

Private Sub AddedDevicesApp_ReadProcess(readData As String)
On Error GoTo er:
'    AddedDevicesApp.Cancel
    readData = Trim(readData)
   ' MsgBox readData
   ' readData = Replace(readData, vbNewLine, ";")
    
    If Left(readData, 1) = ";" Then
        readData = Mid(readData, 2)
    End If
'    Clipboard.Clear
'    Clipboard.SetText readData
    Dim yy() As String
    yy = Split(readData, vbNewLine)
    Dim strDev As String
    Dim strComputer As String
    Dim unreachable As Boolean
    Dim i As Integer
    Dim LastComputer As String
    For i = 0 To UBound(yy)
        strDev = yy(i)
        If Trim(strDev) = "" Then GoTo nxt:
        If InStr(1, strDev, vbTab) < 1 And strDev <> "" Then
            strComputer = Mid(strDev, 1, InStr(1, strDev, ":") - 1)
            LastComputer = strComputer
        Else
            strDev = Replace(strDev, vbTab, ""): strDev = Trim(strDev)
            If InStr(1, strComputer, "[") Then strComputer = LastComputer
            strComputer = strDev & " [" & strComputer + "]"
            Dim itmx As ListItem
            Set itmx = ListView1.ListItems.Add(, , strComputer, 1, 1)
            itmx.Tag = Trim(Mid(strDev, InStr(1, strDev, "- ") + 2))
            If InStr(1, itmx.Tag, "(") Then itmx.Tag = Mid(itmx.Tag, 1, InStr(1, itmx.Tag, "(") - 1)
            If unreachable = False Then
                itmx.Tag = Trim(itmx.Tag)
            Else
                itmx.Tag = ""
            End If
            
        End If
nxt:
    Next
    

    
        
    Exit Sub
er:
    Screen.MousePointer = vbDefault
    Label1(1).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(1).ForeColor = vbRed

End Sub

Private Sub ConnectDevicesApp_ReadProcess(readData As String)
    On Error GoTo er:
    readData = Trim(readData)
    ConnectDevicesApp.Cancel
    readData = Replace(readData, vbNewLine, ";")
    If InStr(1, readData, "ERROR") Then
       ' Label1(1).Caption = "Error: " & vbNewLine & "" & "USB Device Is not Pluged in."
        MsgBox "Error: " & vbNewLine & "" & "USB Device Is not Pluged in.", vbCritical
        Exit Sub
    End If
    If InStr(1, readData, "successfully") Then
'        Winsock2.LocalPort = 0
'        Winsock2.Connect GetSelectedPCName, 33002
        Exit Sub
    End If
    If Strings.Left(readData, 1) = ";" Then
        lvButtons_H1(4).Enabled = Not MakeConnectOperation
        lvButtons_H1(5).Enabled = MakeConnectOperation
        SaveSetting App.CompanyName, "RemoteDevices", ListView1.SelectedItem.Tag, IIf(MakeConnectOperation, 1, 0)
    Else
        lvButtons_H1(5).Enabled = Not MakeConnectOperation
        lvButtons_H1(4).Enabled = MakeConnectOperation
    End If

        
    Exit Sub
er:
    Screen.MousePointer = vbDefault
    Label1(1).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(1).ForeColor = vbRed
End Sub

Private Sub Form_Initialize()
    InitCommonControlsXP
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseMove Button, Shift, x, y
End Sub

Private Sub LoadDevicesApp_ReadProcess(readData As String)
    On Error GoTo er:
    
    readData = Trim(readData)
    readData = Replace(readData, vbNewLine, ";")
    If Left(readData, 1) = ";" Then
        readData = Mid(readData, 2)
    End If
    
    Dim Name As String
    If Left(readData, 3) = "USB" Then
        
        Name = Mid(readData, InStr(1, readData, "- ") + 5)
        Name = Mid(Name, 1, InStr(1, Name, ";") - 1)
        Name = Trim(Name)
        Dim itmx As ListItem
        Set itmx = ListView1.ListItems.Add(, , Name, 1, 1)
        itmx.Tag = Trim(Mid(readData, 1, InStr(1, readData, "- ") - 1))
        
    End If
        
    Exit Sub
er:
    Screen.MousePointer = vbDefault
    Label1(1).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(1).ForeColor = vbRed
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove 0, 0, 0, 0
End Sub
Private Sub Form_Load()
    Image1.Picture = Me.Picture
    Image1.Top = 0
    Image1.Left = 0
    Image1.Width = Me.ScaleWidth
    Image1.Height = Me.ScaleHeight

    If CheckVer(Label1(0), False) = False Then
        MsgBox "Your trail period is expired." + vbNewLine + "Please contact with:Ashrafnet4u@hotmail.com" + vbNewLine + "Desktop Team", vbCritical
        End
    Else
        Set LoadDevicesApp = New ConsoleCapture
        Set AddedDevicesApp = New ConsoleCapture
        Set ConnectDevicesApp = New ConsoleCapture
        
        SetForm Me, &HFF00FF
        lblNav_Click 0
        If ListView1.ListItems.Count = 0 Then
            lblNav_Click 1
        End If
        Timer1.Enabled = True
        On Error Resume Next
        Winsock1.Listen
'        Label1(0) = "Winsock1.Listen"
    End If
    

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblNav(0).FontUnderline = False
    lblNav(1).FontUnderline = False
    lblNav(2).FontUnderline = False
    lblNav(2).ForeColor = &H8000&
    lblNav(0).ForeColor = vbBlue
    lblNav(1).ForeColor = vbBlue
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.Hwnd, &HA1, 2, 0&
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown 0, 0, 0, 0
End Sub

Private Sub lblNav_Click(Index As Integer)
    Dim isvisiable As Boolean
    
    If Index <> 2 Then
     lblNav(0).Enabled = True: lblNav(0).FontItalic = False
    lblNav(1).Enabled = True: lblNav(1).FontItalic = False
    lvButtons_H1(4).Enabled = False
    lvButtons_H1(5).Enabled = False
    lvButtons_H1(6).Enabled = False
    End If
    Select Case Index
        Case 0
            lblNav(2).Visible = True: lvButtons_H1(6).Visible = True
            lblNav(Index).Enabled = False: lblNav(Index).FontItalic = True
            isvisiable = False
        Case 1
            lblNav(2).Visible = False: lvButtons_H1(6).Visible = False
            isvisiable = True
            lblNav(Index).Enabled = False: lblNav(Index).FontItalic = True
            ListView1.ListItems.Clear
            Label1(1) = "Ready, Type Remote IP address then press return."
            Label1(1).ForeColor = vbBlue
            
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1)
        Case 2 ' refresh
            If lblNav(0).Enabled = False Then 'refresh recently devices
                LoadAddedDevices
            Else ' refresh remote devices
                lvButtons_H1_Click 1
            End If
            Exit Sub
    End Select
    
   
    
    Label1(2).Visible = isvisiable
    Text1.Visible = isvisiable
    lvButtons_H1(1).Visible = isvisiable
    
    If isvisiable Then
        ListView1.Height = 1455
        ListView1.Top = 1440
        Text1.SetFocus
    Else
        ListView1.Height = 2055
        ListView1.Top = 840
        DoEvents
        LoadAddedDevices
        
        
    End If
    Label1(1).Top = ListView1.Top
End Sub

Private Sub lblNav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not lblNav(Index).FontUnderline Then
        lblNav(Index).FontUnderline = True
        lblNav(Index).ForeColor = vbRed
    End If
End Sub
Sub SetStatus(strMsg As String)
    Label1(1).Caption = strMsg
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo er:
    
    Dim State As String
    State = 0
        If Item.Tag = "" Then
            State = 2
        ElseIf InStr(1, Item.Text, "(Connected)") Then
            State = 1
        End If
    lvButtons_H1(4).Enabled = False
    lvButtons_H1(5).Enabled = False
    lvButtons_H1(6).Enabled = False
    Label1(1).Caption = Item.Text
    Label1(1).FontBold = True
    
    If State = 0 Then
        SetStatus Item.Text & vbNewLine & "Not Connected yet."
        Label1(1).ForeColor = vbBlack
        lvButtons_H1(4).Enabled = True
        lvButtons_H1(6).Enabled = True
        Item.SmallIcon = 1
    ElseIf State = 1 Then  ' Connected
        SetStatus Item.Text & ""
        Label1(1).ForeColor = &H8000&
        Label1(1).FontBold = True
        lvButtons_H1(5).Enabled = True
        Item.SmallIcon = 3
    ElseIf State = 2 Then
        SetStatus Item.Text & " [Unreachable]"
        Label1(1).ForeColor = vbRed
        Label1(1).FontBold = True
        lvButtons_H1(6).Enabled = True
        
        Item.SmallIcon = 4
    End If
    Exit Sub
er:
    Label1(1).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(1).ForeColor = vbRed
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)
    Select Case Index
        Case 0 'Close App
            Unload Me
            Exit Sub
        Case 1 ' Find
            lvButtons_H1(1).Enabled = False
            StartFinding Trim(Text1.Text)
            lvButtons_H1(1).Enabled = True
        Case 2 ' About form
            frmAbout.Show , Me
        Case 3 'register
            frmReg.Show , Me
        Case 4 ' Connect
            If ListView1.ListItems.Count < 1 Then Exit Sub
            ConnectDevice True
            
        Case 5 ' Disconnect
            If ListView1.ListItems.Count < 1 Then Exit Sub
           ConnectDevice False
            
        Case 6 ' Remove Added Device
            RemoveDevice
            
    End Select
    If ListView1.ListItems.Count = 0 Then
        lvButtons_H1(4).Enabled = False
        lvButtons_H1(5).Enabled = False
        lvButtons_H1(6).Enabled = False
    End If
End Sub

Sub ConnectDevice(MakeConnect As Boolean)
On Error GoTo er:
    
   If ListView1.ListItems.Count < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    MakeConnectOperation = MakeConnect
    
    Dim curCmd As String
    Dim strShare As String
    strShare = "connect"
    If MakeConnect = False Then strShare = "disconnect"
    curCmd = GetShortName(App.Path & "\USBCLNCmd.dat") & " " & strShare & " " & GetSelectedPCName & ":33000:" & ListView1.SelectedItem.Tag
   'Clipboard.Clear: Clipboard.SetText curCmd
    ConnectDevicesApp.RunProcess "cmd /k " & curCmd     'Launch the console
    
    LoadAddedDevices
    'ListView1_ItemClick ListView1.SelectedItem
    Screen.MousePointer = vbDefault
    Exit Sub
er:

    Label1(2).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(2).ForeColor = vbRed
    Screen.MousePointer = vbDefault
End Sub

Function GetSelectedPCName()
    Dim PC As String
    PC = ListView1.SelectedItem.Text
    PC = Mid(PC, InStr(1, PC, "[") + 1)
    PC = Mid(PC, 1, InStr(1, PC, "]") - 1)
    GetSelectedPCName = PC
End Function
Sub RemoveDevice()
On Error GoTo er:
    
   If ListView1.ListItems.Count < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    
    Dim curCmd As String
    Dim PC As String
    PC = ListView1.SelectedItem.Text
   ' MsgBox InStr(1, PC, "[")
    PC = Mid(PC, InStr(1, PC, "[") + 1)
    PC = Mid(PC, 1, InStr(1, PC, "]") - 1)
    Dim strShare As String
    strShare = "del"
    
    curCmd = GetShortName(App.Path & "\USBCLNCmd.dat") & " " & strShare & " " & PC & " 33000"
   'Clipboard.Clear: Clipboard.SetText curCmd
    ConnectDevicesApp.RunProcess "cmd /k " & curCmd     'Launch the console
    
    LoadAddedDevices
    'ListView1_ItemClick ListView1.SelectedItem
    Screen.MousePointer = vbDefault
    Exit Sub
er:

    Label1(2).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(2).ForeColor = vbRed
    Screen.MousePointer = vbDefault
End Sub

Sub StartFinding(TargetHost As String)
    Dim Index As Long
    Dim Name As Variant
    
    
    ListView1.ListItems.Clear

    Label1(1).Caption = "Looking on '" + TargetHost + "' for shared Usb devices..."
    Screen.MousePointer = vbHourglass
    
    Dim curCmd As String
    curCmd = GetShortName(App.Path & "\USBCLNCmd.dat") & " add " & TargetHost & " 33000"
    LoadDevicesApp.RunProcess "cmd /k " & curCmd
    Sleep 2000
    LoadAddedDevices
    If ListView1.ListItems.Count > 0 Then ListView1_ItemClick ListView1.SelectedItem
    Label1(1).Caption = "Available USB Devices" & vbNewLine & "[" & ListView1.ListItems.Count & " Device/s]"
    Screen.MousePointer = vbDefault
End Sub

Sub LoadAddedDevices()

    ListView1.ListItems.Clear
    Label1(1).Caption = "Loading Recently Usb devices..."
    Screen.MousePointer = vbHourglass
    
    Dim curCmd As String
    
    curCmd = GetShortName(App.Path & "\USBCLNCmd.dat") & " list -a"
    AddedDevicesApp.RunProcess "cmd /k " & curCmd     'Launch the console
    chechStates
    If ListView1.ListItems.Count > 0 Then
        ListView1_ItemClick ListView1.SelectedItem
    Else
        lvButtons_H1(4).Enabled = False
        lvButtons_H1(5).Enabled = False
        lvButtons_H1(6).Enabled = False
    End If
    Label1(1).Caption = "Available USB Devices" & vbNewLine & "[" & ListView1.ListItems.Count & " Device/s]"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Text1_Change()
    If Len(Trim(Text1.Text)) > 0 Then
        lvButtons_H1(1).Enabled = True
    Else
        lvButtons_H1(1).Enabled = False
    End If
End Sub

Private Sub Text1_GotFocus()
    lvButtons_H1(1).Default = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    chechStates
    On Error Resume Next
'    Label1(0) = Winsock1.State
    If Winsock1.State = 7 Then
        Winsock1.Close
        Winsock1.Listen
    End If
    Timer1.Enabled = True
End Sub

Sub chechStates()
        On Error GoTo er:
    Dim i As Integer
    Dim State As String
    
    For i = 1 To ListView1.ListItems.Count
        State = 0
        If ListView1.ListItems(i).Tag = "" Then
            State = 2
        ElseIf InStr(1, ListView1.ListItems(i).Text, "(Connected)") Then
            State = 1
        End If
        If State = 0 Then
            ListView1.ListItems(i).SmallIcon = 1
        ElseIf State = 1 Then
            ListView1.ListItems(i).SmallIcon = 3
        ElseIf State = 2 Then
            ListView1.ListItems(i).SmallIcon = 4
        End If
    Next i
    If ListView1.ListItems.Count > 0 Then ListView1_ItemClick ListView1.SelectedItem
    Exit Sub
er:
    Resume
    Label1(1).Caption = "Error: " & vbNewLine & "" & Err.Description
    Label1(1).ForeColor = vbRed

End Sub


Private Sub Winsock1_Connect()
'    Label1(0) = "Connected_" & Winsock1.State
On Error Resume Next
'    Label1(0) = "Winsock1_Connect"
    Winsock1.Close
    Winsock1.Listen
End Sub
Private Sub Winsock1_Close()
    On Error Resume Next
'    Label1(0) = "Winsock1_Close"
    Winsock1.Close
    Winsock1.Listen
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'Label1(0) = "Winsock1_ConnectionRequest"
    Winsock1.Close
    Winsock1.Accept requestID
'    Label1(0) = "ConnectionRequest_" & Winsock1.State
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Label1(0) = "Winsock1_DataArrival"
    Dim s As String
    Winsock1.GetData s
    StartFinding Trim(s)
    
    Label1(1) = "New Customer:" & vbNewLine & "[" & s & "]"
    Winsock1.Close
    Winsock1.Listen
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
        On Error Resume Next
        Label1(1) = "Source=" & Source & "_Desc= " & Description & ""
'    Label1(0) = "Winsock1_Close"
    Winsock1.Close
    Winsock1.Listen
End Sub

Private Sub Winsock2_Connect()
On Error Resume Next
    Winsock2.SendData ListView1.SelectedItem.Tag & ";" & MakeConnectOperation
End Sub

