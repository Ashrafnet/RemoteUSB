Attribute VB_Name = "modDLL"
Option Explicit

Public Type FT_ERROR_STATE
    dwLastError As Integer
    unLine As Integer
    szModule As String
    szDescription As String
End Type

Public Type FT_USB_UNIQID
    idVendor As Integer
    idProduct As Integer
    bcdDevice As Integer
    szSerialNumber As String
End Type

Public Enum eFtUsbDeviceStatus
    eFtUsbDeviceNotShared
    eFtUsbDeviceSharedActive
    eFtUsbDeviceSharedNotActive
    eFtUsbDeviceSharedNotPlugged
    eFtUsbDeviceSharedProblem
End Enum

Public Type FT_SERVER_USB_DEVICE
     usbHWID As FT_USB_UNIQID
     status As eFtUsbDeviceStatus
     bExcludeDevice As Integer
     bSharedManually As Integer
     ulDeviceId As Integer
     ulClientAddr As Integer
     szUsbDeviceDescr As String
     szLocationInfo As String
     szNickName As String
End Type
Declare Function FtEnumDevices Lib "ftusbsrv.dll" (lpUsbDevices As FT_SERVER_USB_DEVICE, ByVal pulBufferSize As Integer, lpES As FT_ERROR_STATE) As Integer
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpTo As Any, lpFrom As Any, ByVal lLen As Long)

Public m_fterror As FT_ERROR_STATE
Public m_pUsbDevs As FT_SERVER_USB_DEVICE

Sub loaddevices()
    Dim ulDeviceBufferSize As Long
    If FtEnumDevices(m_pUsbDevs, ulDeviceBufferSize, m_fterror) = 0 Then
            MsgBox Err.Description, vbCritical
            
        End If
'
'    Do
'        ' Get buffer size needed to enumbrate devices
'        If FtEnumDevices(m_pUsbDevs, ulDeviceBufferSize, m_fterror) = 0 Then
'            MsgBox Err.Description, vbCritical
'            Exit Do
'        End If
'
'        ' No devices
'        If ulDeviceBufferSize = 0 Then
'            Exit Do
'        End If
'
'
'        CopyMem m_pUsbDevs, ulDeviceBufferSize, ulDeviceBufferSize
'
'
'        If IsNull(m_pUsbDevs) Then
'            Exit Do
'        End If
'
'        ' Enumerate devices
'        If FtEnumDevices(m_pUsbDevs, ulDeviceBufferSize, m_fterror) = 0 Then
'            MsgBox Err.Description, vbCritical
'            Exit Do
'        End If
'
'        ' Fill treeview
'        For i = 0 To ulDeviceBufferSize / sizeof(FT_SERVER_USB_DEVICE) - 1
'           '' s.Format(IDS_USBDEVICE_HWID, m_pUsbDevs(i).usbHWID.idVendor, m_pUsbDevs(i).usbHWID.idProduct)
'
'            hItem = m_devicetree.InsertItem(s, hRoot)
'          ''  m_devicetree.SetItemData(hItem, (UInteger) And m_pUsbDevs(i))
'        Next i
'
'    Loop While 0

End Sub
