VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Text1.LinkMode = 0
 Text1.LinkTopic = "Project1|MyTopic"
 Text1.LinkItem = "txtSource"
 Text1.LinkTimeout = 100
 Text1.LinkMode = vbLinkAutomatic
 
End Sub
