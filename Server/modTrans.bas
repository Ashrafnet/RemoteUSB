Attribute VB_Name = "modTrans"
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Type POINTAPI
   x As Long
   y As Long
End Type


Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const LWA_COLORKEY = &H1

Sub SetForm(frm As Form, Optional Color As Long = vbBlack)
  Dim Ret As Long
  Dim CLR As Long
  frm.Hide

  'CLR = RGB(0, 0, 0)  'this color is the color that will be transparent
  CLR = Color
  'Set the window style to 'Layered'
  Ret = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
  Ret = Ret Or WS_EX_LAYERED
  SetWindowLong frm.hwnd, GWL_EXSTYLE, Ret
  'Set the opacity of the layered window to 128
  SetLayeredWindowAttributes frm.hwnd, CLR, 0, LWA_COLORKEY
'  OpenFile
'  SetTopMost
  'SetStartUp
  frm.Show

End Sub
