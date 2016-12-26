Attribute VB_Name = "modThread"
Option Explicit
Public Declare Function CreateThread Lib "kernel32" _
(ByVal lpThreadAttributes As Any, ByVal dwStackSize As _
Long, ByVal lpStartAddress As Long, lpParameter As _
Any, ByVal dwCreationFlags As Long, lpThreadID As _
Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal _
dwMilliseconds As Long)
Public Declare Function TerminateThread Lib "kernel32" _
(ByVal hThread As Long, ByVal dwExitCode As Long) As _
Long
Public Declare Function TerminateProcess Lib _
"kernel32" (ByVal hProcess As Long, ByVal uExitCode As _
Long) As Long
Public Declare Function GetCurrentProcess Lib _
"kernel32" () As Long

Public Declare Function WaitForSingleObject Lib _
"kernel32.dll" (ByVal hHandle As Long, ByVal _
dwMilliseconds As Long) As Long
Public Declare Function CreateEvent& Lib "kernel32" _
Alias "CreateEventA" (ByVal lpEventAttributes As Long, _
ByVal bManualReset As Long, ByVal bInitialState As _
Long, ByVal lpname As String)

Public lThreadHandle1 As Long
Public lThreadHandle2 As Long
Public lEventHandle As Long

Public Sub test_function()
    Dim i As Long
    Dim ret As Long
    
    i = 0
    
    
    Do
        
    Loop Until ret = 0


End Sub


