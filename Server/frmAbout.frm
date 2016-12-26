VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   1845
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Ok"
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
      Image           =   "frmAbout.frx":117E
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":15D1
      ForeColor       =   &H00FF0000&
      Height          =   555
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "We Are the leaders in our field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   795
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Desktop Team]"
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
      Index           =   1
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   4320
      Picture         =   "frmAbout.frx":1675
      Top             =   360
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remote USB-About US"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown 0, 0, 0, 0
End Sub


Private Sub Form_Load()
    Image2.Picture = Me.Picture
    Image2.Top = 0
    Image2.Left = 0
    Image2.Width = Me.ScaleWidth
    Image2.Height = Me.ScaleHeight
    SetForm Me, &HFF00FF
    Label1(0).Caption = "Amerizon Wireless -About US"
    Label2(2).Caption = "My IP Address:"
    
    If Len(Trim(frmVPN.VPNIP)) > 0 Then
        Label2(2).Caption = Label2(2).Caption + vbNewLine + frmVPN.VPNIP
        Exit Sub
    End If
    Dim IpAddrs
    IpAddrs = GetIpAddrTable
    Dim i As Integer
    For i = LBound(IpAddrs) To UBound(IpAddrs)
        If InStr(1, IpAddrs(i), "127.0.") < 1 Then
            Label2(2).Caption = Label2(2).Caption + vbNewLine + IpAddrs(i)
        End If
    Next
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseDown 0, 0, 0, 0
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)
    Select Case Index
        Case 0 'Close App
            Unload Me
    End Select
End Sub


