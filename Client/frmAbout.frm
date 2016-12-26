VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   1860
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RemoteUSBClient.lvButtons_H lvButtons_H1 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amerizon Wireless -About US"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2475
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
      Height          =   555
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "One of Desktop Team Products."
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
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   4200
      Picture         =   "frmAbout.frx":15D1
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown 0, 0, 0, 0
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown 0, 0, 0, 0
End Sub

Private Sub Form_Load()
SetForm Me, &HFF00FF
    Image2.Picture = Me.Picture
    Image2.Top = 0
    Image2.Left = 0
    Image2.Width = Me.ScaleWidth
    Image2.Height = Me.ScaleHeight
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.Hwnd, &HA1, 2, 0&
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


