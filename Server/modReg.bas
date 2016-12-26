Attribute VB_Name = "modReg"
Option Explicit


'=============================================================================================================
'
' modRegistry Module
' ------------------
'
' Last Update : April 01, 2000
'
' VB Versions : 5.0 / 6.0
'
' Requires    : NOTHING
'
' Description : This module was created to easily read from / write to the
'               Windows registry.
'
' Example Use :
'
'  SaveString HKEY_CURRENT_USER, "Software/! Test", "Testing", "Hello World!"
'  MsgBox GetString(HKEY_CURRENT_USER, "Software/! Test", "Testing")
'
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


Public Enum RegistryKeys
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Public Sub SaveKey(ByVal hKey As RegistryKeys, ByVal strPath As String)
On Error Resume Next
  
  Dim KeyHand As Long
  
  RegCreateKey hKey, strPath, KeyHand
  RegCloseKey KeyHand
  
End Sub

Public Function DeleteKey(ByVal hKey As RegistryKeys, ByVal strKey As String)
On Error Resume Next
  
  RegDeleteKey hKey, strKey

End Function

Public Function DeleteValue(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String)
On Error Resume Next

  Dim KeyHand As Long
  
  RegOpenKey hKey, strPath, KeyHand
  RegDeleteValue KeyHand, strValue
  RegCloseKey KeyHand

End Function
' y=GetString (HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control","{N072C50DE872A3F737C8}")
Public Function GetString(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String) As String
On Error Resume Next

  Dim KeyHand As Long
  Dim datatype As Long
  Dim lResult As Long
  Dim strBuf As String
  Dim lDataBufSize As Long
  Dim intZeroPos As Integer
  Dim lValueType As Long
  
  RegOpenKey hKey, strPath, KeyHand
  lResult = RegQueryValueEx(KeyHand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
  If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(KeyHand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
      intZeroPos = InStr(strBuf, Chr(0))
      If intZeroPos > 0 Then
        GetString = Left(strBuf, intZeroPos - 1)
      Else
        GetString = strBuf
      End If
    End If
  End If
    
End Function

Public Sub SaveString(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String, ByVal strData As String)
On Error Resume Next

  Dim KeyHand As Long
  
  RegCreateKey hKey, strPath, KeyHand
  RegSetValueEx KeyHand, strValue, 0, REG_SZ, ByVal strData, Len(strData)
  RegCloseKey KeyHand

End Sub

Function GetDWORD(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValueName As String) As Long
On Error Resume Next

  Dim lResult As Long
  Dim lValueType As Long
  Dim lBuf As Long
  Dim lDataBufSize As Long
  Dim KeyHand As Long

  RegOpenKey hKey, strPath, KeyHand
  lDataBufSize = 4
  lResult = RegQueryValueEx(KeyHand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

  If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
      GetDWORD = lBuf
    End If
  End If

  RegCloseKey KeyHand
    
End Function

Function SaveDWORD(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
On Error Resume Next

   Dim lResult As Long
   Dim KeyHand As Long
   
   RegCreateKey hKey, strPath, KeyHand
   lResult = RegSetValueEx(KeyHand, strValueName, 0&, REG_DWORD, lData, 4)
   RegCloseKey KeyHand
    
End Function


