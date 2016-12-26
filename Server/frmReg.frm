VERSION 5.00
Begin VB.Form frmReg 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   Picture         =   "frmReg.frx":0000
   ScaleHeight     =   1860
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Cancel"
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
      Image           =   "frmReg.frx":117E
      cBack           =   -2147483633
   End
   Begin RemoteUSBServer.lvButtons_H lvButtons_H1 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   2
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
      Image           =   "frmReg.frx":13F5
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial"
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
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remote USB-Registration"
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
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents conApp As ConsoleCapture
Attribute conApp.VB_VarHelpID = -1
Dim bLocked As Boolean      'Locked?



Private Sub Form_Load()

    SetForm Me, &HFF00FF
    
     Set conApp = New ConsoleCapture             'Initialise conApp object
    'conApp.RunProcess "cmd /k", vbNullString    'Load the command prompt once
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub lvButtons_H1_Click(Index As Integer)
    Select Case Index
        Case 0 'register
            If bLocked = True Then Exit Sub
            If Len(Trim(Text1(0))) = 0 Then MsgBox "Please, Enter your registration's Name.", vbCritical: Text1(0).SetFocus: Exit Sub
            If Len(Trim(Text1(1))) = 0 Then MsgBox "Please, Enter your registration's Serial.", vbCritical: Text1(1).SetFocus: Exit Sub
            lvButtons_H1(0).Enabled = False
            lvButtons_H1(1).Enabled = False
            Screen.MousePointer = vbHourglass
            Dim curCmd As String
            Dim servicename As String
            bLocked = True                                  'Lock the output window
            If IsFileExist(App.Path & "\UsbService64.exe") Then
                servicename = "UsbService64.exe"
            ElseIf IsFileExist(App.Path & "\UsbService.exe") Then
                servicename = "UsbService.exe"
            Else
                MsgBox "The UsbService.exe is not exist." + vbNewLine + "Please reinstall the Remote Usb Software, then try again.", vbCritical
                GoTo er:
            End If
            curCmd = "" & servicename & " REG " & Chr(34) & Text1(0).Text & Chr(34) & " " & Chr(34) & Text1(1).Text & Chr(34)
            conApp.RunProcess "cmd /k " & curCmd, App.Path     'Launch the console
            bLocked = False                                 'Unlock
                                                
            
        Case 1 'Close App
            Unload Me
            Exit Sub
    End Select
er:
    Screen.MousePointer = vbDefault
    lvButtons_H1(0).Enabled = True
            lvButtons_H1(1).Enabled = True
            bLocked = False
End Sub

Function IsFileExist(FilePath As String) As Boolean
    If Dir(FilePath) <> "" Then
        IsFileExist = True
    Else
        IsFileExist = False
    End If
End Function



Private Sub Form_Unload(Cancel As Integer)
    Set conApp = Nothing                        'Unload conApp
End Sub

Private Sub conApp_ReadProcess(readData As String)
    On Error GoTo er:
    'Cancel console application
    conApp.Cancel
    
    If InStr(1, LCase(readData), LCase("Thank you for registration")) > 0 Then
        
        MsgBox "Thank you for registration!", vbInformation
        Screen.MousePointer = vbDefault
        Unload Me
        Exit Sub
    Else
        MsgBox "Your registration Key is not valid!", vbCritical
        
    End If
er:
    lvButtons_H1(0).Enabled = True
    lvButtons_H1(1).Enabled = True

    Screen.MousePointer = vbDefault
End Sub


