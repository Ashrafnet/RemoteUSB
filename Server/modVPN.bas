Attribute VB_Name = "modVPN"
Option Explicit
Public Type RASIPADDR
a As Byte
b As Byte
c As Byte
d As Byte
End Type

Public Enum RasEntryOptions
    RASEO_UseCountryAndAreaCodes = &H1
    RASEO_SpecificIpAddr = &H2
    RASEO_SpecificNameServers = &H4
    RASEO_IpHeaderCompression = &H8
    RASEO_RemoteDefaultGateway = &H10
    RASEO_DisableLcpExtensions = &H20
    RASEO_TerminalBeforeDial = &H40
    RASEO_TerminalAfterDial = &H80
    RASEO_ModemLights = &H100
    RASEO_SwCompression = &H200
    RASEO_RequireEncryptedPw = &H400
    RASEO_RequireMsEncryptedPw = &H800
    RASEO_RequireDataEncryption = &H1000
    RASEO_NetworkLogon = &H2000
    RASEO_UseLogonCredentials = &H4000
    RASEO_PromoteAlternates = &H8000
    RASEO_SecureLocalFiles = &H10000
    RASEO_RequireEAP = &H20000
    RASEO_RequirePAP = &H40000
    RASEO_RequireSPAP = &H80000
    RASEO_Custom = &H100000
    RASEO_PreviewPhoneNumber = &H200000
    RASEO_SharedPhoneNumbers = &H800000
    RASEO_PreviewUserPw = &H1000000
    RASEO_PreviewDomain = &H2000000
    RASEO_ShowDialingProgress = &H4000000
    RASEO_RequireCHAP = &H8000000
    RASEO_RequireMsCHAP = &H10000000
    RASEO_RequireMsCHAP2 = &H20000000
    RASEO_RequireW95MSCHAP = &H40000000
    RASEO_CustomScript = &H80000000
End Enum

Public Enum RASNetProtocols
    RASNP_NetBEUI = &H1
    RASNP_Ipx = &H2
    RASNP_Ip = &H4
End Enum

Public Enum RasFramingProtocols
    RASFP_Ppp = &H1
    RASFP_Slip = &H2
    RASFP_Ras = &H4
End Enum


Public Type VBRasEntry
options As RasEntryOptions
CountryID As Long
CountryCode As Long
AreaCode As String
LocalPhoneNumber As String
AlternateNumbers As String
ipAddr As RASIPADDR
ipAddrDns As RASIPADDR
ipAddrDnsAlt As RASIPADDR
ipAddrWins As RASIPADDR
ipAddrWinsAlt As RASIPADDR
FrameSize As Long
fNetProtocols As RASNetProtocols
FramingProtocol As RasFramingProtocols
ScriptName As String
AutodialDll As String
AutodialFunc As String
DeviceType As String
DeviceName As String
X25PadType As String
X25Address As String
X25Facilities As String
X25UserData As String
Channels As Long
NT4En_SubEntries As Long
NT4En_DialMode As Long
NT4En_DialExtraPercent As Long
NT4En_DialExtraSampleSeconds As Long
NT4En_HangUpExtraPercent As Long
NT4En_HangUpExtraSampleSeconds As Long
NT4En_IdleDisconnectSeconds As Long
Win2000_Type As Long
Win2000_EncryptionType As Long
Win2000_CustomAuthKey As Long
Win2000_guidId(0 To 15) As Byte
Win2000_CustomDialDll As String
Win2000_VpnStrategy As Long
End Type
'Make a combo box for the modem devices and use the GetDevices command.
'in the form Dim clsVbRasEntry As VbRasEntry
'make calls as clsVbRasEntry.options = selected options
'clsVbRasEntry.LocalPhoneNumber = "555-5555" and so forth

Public Declare Function RasSetEntryProperties _
Lib "rasapi32.dll" Alias "RasSetEntryPropertiesA" _
(ByVal lpszPhonebook As String, _
ByVal lpszEntry As String, _
lpRasEntry As Any, _
ByVal dwEntryInfoSize As Long, _
lpbDeviceInfo As Any, _
ByVal dwDeviceInfoSize As Long) _
As Long

Public Declare Function RasGetErrorString _
Lib "rasapi32.dll" Alias "RasGetErrorStringA" _
(ByVal uErrorValue As Long, ByVal lpszErrorString As String, _
cBufSize As Long) As Long

Public Declare Function FormatMessage _
Lib "kernel32" Alias "FormatMessageA" _
(ByVal dwFlags As Long, lpSource As Any, _
ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
ByVal lpBuffer As String, ByVal nSize As Long, _
Arguments As Long) As Long

Public Declare Function RasGetEntryProperties _
Lib "rasapi32.dll" Alias "RasGetEntryPropertiesA" _
(ByVal lpszPhonebook As String, _
ByVal lpszEntry As String, _
lpRasEntry As Any, _
lpdwEntryInfoSize As Long, _
lpbDeviceInfo As Any, _
lpdwDeviceInfoSize As Long) As Long
Public Type VBRASDEVINFO
DeviceType As String
DeviceName As String
End Type

Public Declare Function RasEnumDevices _
Lib "rasapi32.dll" Alias "RasEnumDevicesA" ( _
lpRasDevInfo As Any, _
lpcb As Long, _
lpcDevices As Long _
) As Long



Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

Global Const RAS_MaxDeviceType = 16
Global Const RAS_MaxDeviceName = 128

Global Const GMEM_FIXED = &H0
Global Const GMEM_ZEROINIT = &H40
Global Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Global Const ApINULL = 0&

Type RASDEVINFO
dwSize As Long
szDeviceType(RAS_MaxDeviceType) As Byte
szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Declare Function iRasEnumDevices Lib "rasapi32.dll" Alias "RasEnumDevicesA" ( _
lpRasDevInfo As Any, _
lpcb As Long, _
lpcDevices As Long) As Long

Declare Sub iCopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Sub GetDevices(lst As ComboBox)
Dim lpRasDevInfo As RASDEVINFO
Dim lpcb As Long
Dim cDevices As Long
Dim t_Buff As Long
Dim nRet As Long
Dim t_ptr As Long
Dim i As Long

lpcb = 0

lpRasDevInfo.dwSize = LenB(lpRasDevInfo) + (LenB(lpRasDevInfo) Mod 4)

nRet = iRasEnumDevices(ByVal 0, lpcb, cDevices)

t_Buff = GlobalAlloc(GPTR, lpcb)

iCopyMemory ByVal t_Buff, lpRasDevInfo, LenB(lpRasDevInfo)

nRet = iRasEnumDevices(ByVal t_Buff, lpcb, lpcb)

If nRet = 0 Then
t_ptr = t_Buff

For i = 0 To cDevices - 1
iCopyMemory lpRasDevInfo, ByVal t_ptr, LenB(lpRasDevInfo)
lst.AddItem (ByteToString(lpRasDevInfo.szDeviceName))
t_ptr = t_ptr + LenB(lpRasDevInfo) + (LenB(lpRasDevInfo) Mod 4)
Next i
Else
MsgBox nRet
End If

If t_Buff <> 0 Then GlobalFree (t_Buff)

End Sub

Function ByteToString(bytearray() As Byte) As String
Dim i As Integer, t As String
i = 0
t = ""
While i < UBound(bytearray) And bytearray(i) <> 0
t = t & Chr$(bytearray(i))
i = i + 1
Wend
ByteToString = t
End Function


Function VBRasSetEntryProperties(strEntryName As String, _
typRasEntry As VBRasEntry, _
Optional strPhoneBook As String) As Long

Dim rtn As Long, lngCb As Long, lngBuffLen As Long
Dim b() As Byte
Dim lngPos As Long, lngStrLen As Long

rtn = RasGetEntryProperties(vbNullString, vbNullString, _
ByVal 0&, lngCb, ByVal 0&, ByVal 0&)

If rtn <> 603 Then VBRasSetEntryProperties = rtn: Exit Function

lngStrLen = Len(typRasEntry.AlternateNumbers)
lngBuffLen = lngCb + lngStrLen + 1
ReDim b(lngBuffLen)

CopyMemory b(0), lngCb, 4
CopyMemory b(4), typRasEntry.options, 4
CopyMemory b(8), typRasEntry.CountryID, 4
CopyMemory b(12), typRasEntry.CountryCode, 4
CopyStringToByte b(16), typRasEntry.AreaCode, 11
CopyStringToByte b(27), typRasEntry.LocalPhoneNumber, 129

If lngStrLen > 0 Then
CopyMemory b(lngCb), _
ByVal typRasEntry.AlternateNumbers, lngStrLen
CopyMemory b(156), lngCb, 4
End If

CopyMemory b(160), typRasEntry.ipAddr, 4
CopyMemory b(164), typRasEntry.ipAddrDns, 4
CopyMemory b(168), typRasEntry.ipAddrDnsAlt, 4
CopyMemory b(172), typRasEntry.ipAddrWins, 4
CopyMemory b(176), typRasEntry.ipAddrWinsAlt, 4
CopyMemory b(180), typRasEntry.FrameSize, 4
CopyMemory b(184), typRasEntry.fNetProtocols, 4
CopyMemory b(188), typRasEntry.FramingProtocol, 4
CopyStringToByte b(192), typRasEntry.ScriptName, 260
CopyStringToByte b(452), typRasEntry.AutodialDll, 260
CopyStringToByte b(712), typRasEntry.AutodialFunc, 260
CopyStringToByte b(972), typRasEntry.DeviceType, 17
If lngCb = 1672& Then lngStrLen = 33 Else lngStrLen = 129
CopyStringToByte b(989), typRasEntry.DeviceName, lngStrLen
lngPos = 989 + lngStrLen
CopyStringToByte b(lngPos), typRasEntry.X25PadType, 33
lngPos = lngPos + 33
CopyStringToByte b(lngPos), typRasEntry.X25Address, 201
lngPos = lngPos + 201
CopyStringToByte b(lngPos), typRasEntry.X25Facilities, 201
lngPos = lngPos + 201
CopyStringToByte b(lngPos), typRasEntry.X25UserData, 201
lngPos = lngPos + 203
CopyMemory b(lngPos), typRasEntry.Channels, 4

If lngCb > 1768 Then
CopyMemory b(1768), typRasEntry.NT4En_SubEntries, 4
CopyMemory b(1772), typRasEntry.NT4En_DialMode, 4
CopyMemory b(1776), typRasEntry.NT4En_DialExtraPercent, 4
CopyMemory b(1780), typRasEntry.NT4En_DialExtraSampleSeconds, 4
CopyMemory b(1784), typRasEntry.NT4En_HangUpExtraPercent, 4
CopyMemory b(1788), typRasEntry.NT4En_HangUpExtraSampleSeconds, 4
CopyMemory b(1792), typRasEntry.NT4En_IdleDisconnectSeconds, 4

If lngCb > 1796 Then
CopyMemory b(1796), typRasEntry.Win2000_Type, 4
CopyMemory b(1800), typRasEntry.Win2000_EncryptionType, 4
CopyMemory b(1804), typRasEntry.Win2000_CustomAuthKey, 4
CopyMemory b(1808), typRasEntry.Win2000_guidId(0), 16
CopyStringToByte b(1824), typRasEntry.Win2000_CustomDialDll, 260
CopyMemory b(2084), typRasEntry.Win2000_VpnStrategy, 4
End If

End If

rtn = RasSetEntryProperties(strPhoneBook, strEntryName, _
b(0), lngCb, ByVal 0&, ByVal 0&)

VBRasSetEntryProperties = rtn

End Function


Function VBRASErrorHandler(rtn As Long) As String
    Dim strError As String, i As Long
    strError = String(512, 0)
    If rtn > 600 Then
        RasGetErrorString rtn, strError, 512&
    Else
        FormatMessage &H1000, ByVal 0&, rtn, 0&, strError, 512, ByVal 0&
    End If
    i = InStr(strError, Chr$(0))
    If i > 1 Then VBRASErrorHandler = Left$(strError, i - 1)
End Function

Function VBRasGetEntryProperties(strEntryName As String, _
    typRasEntry As VBRasEntry, _
    Optional strPhoneBook As String) As Long
    
    Dim rtn As Long, lngCb As Long, lngBuffLen As Long
    Dim b() As Byte
    Dim lngPos As Long, lngStrLen As Long
    
    rtn = RasGetEntryProperties(vbNullString, vbNullString, _
    ByVal 0&, lngCb, ByVal 0&, ByVal 0&)
    
    rtn = RasGetEntryProperties(strPhoneBook, strEntryName, _
    ByVal 0&, lngBuffLen, ByVal 0&, ByVal 0&)
    
    If rtn <> 603 Then VBRasGetEntryProperties = rtn: Exit Function
    
    ReDim b(lngBuffLen - 1)
    CopyMemory b(0), lngCb, 4
    
    rtn = RasGetEntryProperties(strPhoneBook, strEntryName, _
    b(0), lngBuffLen, ByVal 0&, ByVal 0&)
    
    VBRasGetEntryProperties = rtn
    If rtn <> 0 Then Exit Function
    
    CopyMemory typRasEntry.options, b(4), 4
    CopyMemory typRasEntry.CountryID, b(8), 4
    CopyMemory typRasEntry.CountryCode, b(12), 4
    CopyByteToTrimmedString typRasEntry.AreaCode, b(16), 11
    CopyByteToTrimmedString typRasEntry.LocalPhoneNumber, b(27), 129
    
    CopyMemory lngPos, b(156), 4
    If lngPos <> 0 Then
        lngStrLen = lngBuffLen - lngPos
        typRasEntry.AlternateNumbers = String(lngStrLen, 0)
        CopyMemory ByVal typRasEntry.AlternateNumbers, _
        b(lngPos), lngStrLen
    End If

    CopyMemory typRasEntry.ipAddr, b(160), 4
    CopyMemory typRasEntry.ipAddrDns, b(164), 4
    CopyMemory typRasEntry.ipAddrDnsAlt, b(168), 4
    CopyMemory typRasEntry.ipAddrWins, b(172), 4
    CopyMemory typRasEntry.ipAddrWinsAlt, b(176), 4
    CopyMemory typRasEntry.FrameSize, b(180), 4
    CopyMemory typRasEntry.fNetProtocols, b(184), 4
    CopyMemory typRasEntry.FramingProtocol, b(188), 4
    CopyByteToTrimmedString typRasEntry.ScriptName, b(192), 260
    CopyByteToTrimmedString typRasEntry.AutodialDll, b(452), 260
    CopyByteToTrimmedString typRasEntry.AutodialFunc, b(712), 260
    CopyByteToTrimmedString typRasEntry.DeviceType, b(972), 17
    If lngCb = 1672& Then lngStrLen = 33 Else lngStrLen = 129
    CopyByteToTrimmedString typRasEntry.DeviceName, b(989), lngStrLen
    lngPos = 989 + lngStrLen
    CopyByteToTrimmedString typRasEntry.X25PadType, b(lngPos), 33
    lngPos = lngPos + 33
    CopyByteToTrimmedString typRasEntry.X25Address, b(lngPos), 201
    lngPos = lngPos + 201
    CopyByteToTrimmedString typRasEntry.X25Facilities, b(lngPos), 201
    lngPos = lngPos + 201
    CopyByteToTrimmedString typRasEntry.X25UserData, b(lngPos), 201
    lngPos = lngPos + 203
    CopyMemory typRasEntry.Channels, b(lngPos), 4

If lngCb > 1768 Then
    CopyMemory typRasEntry.NT4En_SubEntries, b(1768), 4
    CopyMemory typRasEntry.NT4En_DialMode, b(1772), 4
    CopyMemory typRasEntry.NT4En_DialExtraPercent, b(1776), 4
    CopyMemory typRasEntry.NT4En_DialExtraSampleSeconds, b(1780), 4
    CopyMemory typRasEntry.NT4En_HangUpExtraPercent, b(1784), 4
    CopyMemory typRasEntry.NT4En_HangUpExtraSampleSeconds, b(1788), 4
    CopyMemory typRasEntry.NT4En_IdleDisconnectSeconds, b(1792), 4
    
    If lngCb > 1796 Then
        CopyMemory typRasEntry.Win2000_Type, b(1796), 4
        CopyMemory typRasEntry.Win2000_EncryptionType, b(1800), 4
        CopyMemory typRasEntry.Win2000_CustomAuthKey, b(1804), 4
        CopyMemory typRasEntry.Win2000_guidId(0), b(1808), 16
        CopyByteToTrimmedString _
        typRasEntry.Win2000_CustomDialDll, b(1824), 260
        CopyMemory typRasEntry.Win2000_VpnStrategy, b(2084), 4
    End If

End If

End Function



Sub CopyByteToTrimmedString(strToCopyTo As String, _
    bPos As Byte, lngMaxLen As Long)
    Dim strTemp As String, lngLen As Long
    strTemp = String(lngMaxLen + 1, 0)
    CopyMemory ByVal strTemp, bPos, lngMaxLen
    lngLen = InStr(strTemp, Chr$(0)) - 1
    strToCopyTo = Left$(strTemp, lngLen)
End Sub


Sub CopyStringToByte(bPos As Byte, _
    strToCopy As String, lngMaxLen As Long)
    Dim lngLen As Long
    lngLen = Len(strToCopy)
    If lngLen = 0 Then
    Exit Sub
    ElseIf lngLen > lngMaxLen Then
    lngLen = lngMaxLen
    End If
    CopyMemory bPos, ByVal strToCopy, lngLen
End Sub




