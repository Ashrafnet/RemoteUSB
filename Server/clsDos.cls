VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsoleCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ReadProcess(readData As String)

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, ByVal lpBytesRead As Long, lpTotalBytesAvail As Long, ByVal lpBytesLeftThisMessage As Long) As Long

Private Const STARTF_USESTDHANDLES = &H100
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0

Private Const BUFFER_SIZE = 1024    'Change at will

Dim hOutWrite As Long   'Handle the other process writes to
Dim hOutRead As Long    'Handle we read from
Dim bCancel As Boolean  'Cancel flag
Dim hProc As Long       'Handle to process

Dim m_cmdLine As String 'Commandline
Dim m_procID As Long    'ProcessID

'Return the commandline used
Public Property Get CommandLine() As String
    CommandLine = m_cmdLine
End Property
'Return the launched process ID
Public Property Get ProcessID() As Long
    ProcessID = m_procID
End Property

Public Function RunProcess(cmdLine As String, Optional args As String = "", Optional startDir As String = vbNullString) As Long
    Dim saPipe As SECURITY_ATTRIBUTES
    Dim siStart As STARTUPINFO
    Dim piNewProc As PROCESS_INFORMATION
    Dim inBuff As String
    Dim bytesRead As Long
    Dim bytesLeft As Long
    Dim junk As Long
    If Trim(args) <> "" Then cmdLine = cmdLine & " " & args
    'Function-level error catcher/forwarder
    On Error GoTo Final_Cleanup
    
    If (hProc) Then
        'Doh! Already running something
        Err.Raise vbObjectError + 400, "ConsoleCapture", "A process is already running"
        Exit Function
    End If
    
    m_cmdLine = cmdLine
    
    saPipe.nLength = Len(saPipe)    'So WinAPI knows what we're doing
    saPipe.bInheritHandle = -1      'New process can use our pipe handles
    saPipe.lpSecurityDescriptor = 0
    
    'Create the pipe
    If CreatePipe(hOutRead, hOutWrite, saPipe, 0) Then

        siStart.cb = Len(siStart)   'So WinAPI knows what we're doing
        
        siStart.hStdOutput = hOutWrite  'Write to here
        siStart.hStdError = hOutWrite   'Write errors to here
        siStart.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW  'Use our handles and our show method
        siStart.wShowWindow = SW_HIDE   'Dont show (hide)
        
        'Create the process
        If CreateProcess(vbNullString, m_cmdLine, ByVal 0, ByVal 0, -1, 0, ByVal 0, startDir, siStart, piNewProc) Then
        
            m_procID = piNewProc.dwProcessId        'Store the procID
            inBuff = String$(BUFFER_SIZE, Chr$(0))  'Allocate the buffer
            
            bCancel = False
            'Go!
            Do Until bCancel
            DoEvents
                'Do we have data to read?
                If PeekNamedPipe(hOutRead, ByVal 0, 0, 0, bytesLeft, 0) Then
            
                    If (bytesLeft) Then
                        'Yes we do!
                        If bytesLeft > BUFFER_SIZE Then bytesLeft = BUFFER_SIZE
                        
                        'Attempt to read from the pipe
                        If ReadFile(hOutRead, ByVal inBuff, bytesLeft, bytesRead, ByVal 0) Then
                            CloseHandle hOutWrite
                            'Raise the event!
                            RaiseEvent ReadProcess(Left$(inBuff, bytesRead))
                            'Reset the buffer
                            ZeroMemory ByVal inBuff, BUFFER_SIZE
                        Else
                            'Problem occured - stop
                            bCancel = True
                        End If
                
                    Else
                        'Let Windows breathe
                        DoEvents
                        
                    End If
                Else
                
                    'Problem occured - stop
                    bCancel = True
                End If
                DoEvents
            Loop
            
        Else
        
            'Doh! CreateProcess didnt work... raise an error
            Err.Raise vbObjectError + 402, "ConsoleCapture", "An error occured while launching the process"
        End If
            
    Else
    
        'Doh! CreatePipe didnt work... raise an error
        Err.Raise vbObjectError + 401, "ConsoleCapture", "An error occured while creating the pipe"
    End If
    
    'Save the process's exit code
    GetExitCodeProcess piNewProc.hProcess, RunProcess

Final_Cleanup:
    'Clean up anything
    CloseHandle hOutRead
    CloseHandle hOutWrite
    CloseHandle piNewProc.hProcess
    CloseHandle piNewProc.hThread
    
    m_procID = 0
    m_cmdLine = ""
    
    'If a local error was raised, pass it up to the caller
    
    If Err.Number Then Err.Raise Err.Number, Err.Source, Err.Description
'Resume
End Function

Public Sub Cancel()
    'Stop!
    bCancel = True
End Sub

