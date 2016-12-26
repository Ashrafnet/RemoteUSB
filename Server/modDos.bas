Attribute VB_Name = "modDos"
ption Explicit
Private Declare Function AllocConsole _
    Lib "kernel32" () _
    As Long
Private Declare Function FreeConsole _
    Lib "kernel32" () _
    As Long
Private Declare Function GetStdHandle _
    Lib "kernel32" ( _
    ByVal nStdHandle As Long) _
    As Long
Private Declare Function ReadConsole _
    Lib "kernel32" Alias "ReadConsoleA" ( _
    ByVal hConsoleInput As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfCharsToRead As Long, _
    lpNumberOfCharsRead As Long, _
    lpReserved As Any) _
    As Long
Private Declare Function SetConsoleMode _
    Lib "kernel32" ( _
    ByVal hConsoleOutput As Long, _
    dwMode As Long) _
    As Long
Private Declare Function SetConsoleTextAttribute _
    Lib "kernel32" ( _
    ByVal hConsoleOutput As Long, _
    ByVal wAttributes As Long) _
    As Long
Private Declare Function SetConsoleTitle _
    Lib "kernel32" Alias "SetConsoleTitleA" ( _
    ByVal lpConsoleTitle As String) _
    As Long
Private Declare Function WriteConsole _
    Lib "kernel32" Alias "WriteConsoleA" ( _
    ByVal hConsoleOutput As Long, _
    ByVal lpBuffer As Any, _
    ByVal nNumberOfCharsToWrite As Long, _
    lpNumberOfCharsWritten As Long, _
    lpReserved As Any) _
    As Long
'Computer System Information
Private Declare Function GetComputerName _
    Lib "kernel32" Alias "GetComputerNameA" ( _
    ByVal lpBuffer As String, nSize As Long) _
    As Long

Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

'SetConsoleTextAttribute color values
Private Const FOREGROUND_BLUE = &H1
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_RED = &H4
Private Const FOREGROUND_INTENSITY = &H8
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_INTENSITY = &H80
'SetConsoleMode (input)
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
'SetConsoleMode (output)
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
'GetComputerName
Public Const MAX_COMPUTERNAME_LENGTH = 31
' Global Variables
Private hConsoleIn As Long ' console input handle
Private hConsoleOut As Long ' console output handle
Private hConsoleErr As Long ' console error handle

