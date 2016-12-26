Attribute VB_Name = "modFunctions"
Option Explicit
Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Const MAX_COMPUTERNAME_LENGTH As Long = 31

Private Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400
Public RemainsDays As Integer

Public RasProccesHandel As Long



Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

Private Declare Function GetTempPath Lib "kernel32" Alias _
"GetTempPathA" (ByVal nBufferLength As Long, ByVal _
lpBuffer As String) As Long

Const MAX_PATH = 260

' This function uses Windows API GetTempPath to get the temporary folder

Public Function GetTempFolder() As String
    
    Dim sFolder As String ' Name of the folder
    Dim lRet As Long ' Return Value

    sFolder = String(MAX_PATH, 0)
    lRet = GetTempPath(MAX_PATH, sFolder)

    If lRet <> 0 Then
        GetTempFolder = Left(sFolder, InStr(sFolder, Chr(0)) - 1)
    Else
        GetTempFolder = vbNullString
    End If
End Function


Public Function DownloadFile(sSourceUrl As String, _
                             sLocalFile As String) As Boolean

  'Download the file. BINDF_GETNEWESTVERSION forces
  'the API to download from the specified source.
  'Passing 0& as dwReserved causes the locally-cached
  'copy to be downloaded, if available. If the API
  'returns ERROR_SUCCESS (0), DownloadFile returns True.
   DownloadFile = URLDownloadToFile(0&, _
                                    sSourceUrl, _
                                    sLocalFile, _
                                    BINDF_GETNEWESTVERSION, _
                                    0&) = ERROR_SUCCESS

End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
       
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)
    
    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)

End Function

' Returns an array with the local IP addresses (as strings).
' Author: Christian d'Heureuse, www.source-code.biz
Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim rc As Long
   rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   For i = 0 To NrOfEntries - 1
      Dim J As Integer, s As String: s = ""
      For J = 0 To 3: s = s & IIf(J > 0, ".", "") & Buf(4 + i * 24 + J): Next
      IpAddrs(i) = s
      Next
   GetIpAddrTable = IpAddrs
   End Function

Public Function EncodeDecode(input1 As String)
    On Error GoTo er:
    Dim x As Integer
    Dim output1 As String
    For x = 1 To Len(input1)
        If Asc(UCase(Mid(input1, x, 1))) > 64 And Asc(UCase(Mid(input1, x, 1))) < 91 Then
            'note: the next two lines on this page are one
            output1 = output1 + LCase(Chr((((Asc(UCase(Mid(input1, x, 1))) - 65) + 13) Mod 26) + 65))
        Else
            output1 = output1 + Mid(input1, x, 1)
        End If
    Next
    EncodeDecode = output1
    Exit Function
er:
    EncodeDecode = ""
End Function

' Convert a date string into a Date value using a format.
' The format should be "ymd", "mdy", etc.
Public Function ToDate(ByVal date_string As String, ByVal _
    date_format As String) As Date
Dim date_parts() As String
Dim day_part As String
Dim month_part As String
Dim year_part As String

    date_parts = Split(date_string, "-")

    date_format = LCase$(date_format)
    Select Case Mid$(date_format, 1, 1)
        Case "d"
            day_part = date_parts(0)
        Case "m"
            month_part = date_parts(0)
        Case "y"
            year_part = date_parts(0)
    End Select
    Select Case Mid$(date_format, 2, 1)
        Case "d"
            day_part = date_parts(1)
        Case "m"
            month_part = date_parts(1)
        Case "y"
            year_part = date_parts(1)
    End Select
    Select Case Mid$(date_format, 3, 1)
        Case "d"
            day_part = date_parts(2)
        Case "m"
            month_part = date_parts(2)
        Case "y"
            year_part = date_parts(2)
    End Select

    ToDate = CDate(day_part & " " & MonthName(month_part) & _
        ", " & year_part)
End Function

Function CheckCRC(RegVal As String, CRC As Long) As Boolean
    On Error GoTo er:
    Dim i As Integer
    Dim c As String
    c = CRC
    c = Mid(c, 1, Len(c) - 3)
    '36537031
    CRC = CLng(c)
    Dim MyCrC As Long
    For i = 1 To Len(RegVal)
        c = Mid(RegVal, i, 1)
        MyCrC = MyCrC + Asc(c) + 211
    Next i
    If CRC = MyCrC Then
        CheckCRC = True
    Else
        CheckCRC = False
    End If
    Exit Function
er:
    CheckCRC = False
End Function

'021911
Function CheckVer(Label1 As Label, IsServer As Boolean) As Integer
    CheckVer = True
    Exit Function
    On Error GoTo er:
    Dim y As String
    Dim x As String
    Dim CRC As Long
    If IsServer Then
        y = GetString(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control", "{N172C5T4TYNHGE272A3F7887HJMNGC173N4}")
    Else
        y = GetString(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control", "{C172C5T4TYNHGE276589RGHBVE654GC173NE}")
    End If
    x = Mid(y, 1, InStr(1, y, ",") - 1)
    CRC = CLng(Mid(y, InStr(1, y, ",") + 1))
    If CheckCRC(x, CRC) = False Then
        CheckVer = False
        Exit Function
    End If
    y = EncodeDecode(y)
    If Len(Trim(y)) < 1 Then
        CheckVer = False
        Exit Function
    End If

    y = Mid(y, InStr(1, y, " - ") + 3)
    Dim myday As String
    Dim mymonth As String
    Dim myyear As String
    Dim mydate As Date
    myday = Mid(y, 3, 2)
    mymonth = Mid(y, 1, 2)
    myyear = Mid(y, 5, 2)
    
   
    If CInt(myyear) > CInt(Mid(Year(Now), 3)) Then
        myyear = "19" & myyear
    Else
        myyear = "20" & myyear
    End If
    mydate = ToDate(myday & "-" & mymonth & "-" & myyear, "dmy")
    
    x = DateDiff("d", mydate, Now)
    RemainsDays = 12 - x
    Label1.Caption = Label1.Caption & " [" & RemainsDays & " Days Remains]"
    If RemainsDays <= 0 Then
        CheckVer = False
    Else
        CheckVer = True
    End If
    Exit Function
er:
    CheckVer = False
End Function

Public Function ShellandWait(ExeFullPath As String, Optional TimeOutValue As Long = 0) As Boolean
    
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessId As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean
    
    On Error GoTo ErrorHandler

    lStart = CLng(Timer)
    sExeName = ExeFullPath

    'Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    lInst = Shell(sExeName, vbHide)
    
    lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst)
    RasProccesHandel = lProcessId
    

    Do
        Call GetExitCodeProcess(lProcessId, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < lStart Then Exit Do
            Else
                 Exit Do
            End If
    End If
    DoEvents
    Loop While lExitCode = STATUS_PENDING
    
    ShellandWait = True
   
ErrorHandler:
ShellandWait = False
Exit Function
End Function

Function ComputerName() As String
    Dim dwLen As Long
    Dim strString As String
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    ComputerName = strString
End Function


