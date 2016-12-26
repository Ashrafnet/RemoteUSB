Attribute VB_Name = "u2ec"
' enum hub
Declare Function ServerCreateEnumUsbDev Lib "RemoteUSB.dll" (ByRef Context As Long) As Boolean
Declare Function ServerRemoveEnumUsbDev Lib "RemoteUSB.dll" (ByVal Context As Long) As Boolean
Declare Function ServerGetUsbDevFromHub Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal HubContext As Long, ByVal Index As Long, ByRef DevContext As Long) As Boolean
Declare Function ServerUsbDevIsHub Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long) As Boolean
Declare Function ServerUsbDevIsShared Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long) As Boolean
Declare Function ServerUsbDevIsConnected Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long) As Boolean
Declare Function ServerGetUsbDevName Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef Name As Variant) As Boolean

Declare Function SetCallBackOnChangeDevList Lib "RemoteUSB.dll" (ByRef Callback As Long) As Boolean
 
' server
Declare Function ServerShareUsbDev Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long, ByVal Connectionn As Variant, ByVal Description As Variant, ByVal Auth As Boolean, ByVal Passw As Variant, ByVal Crypt As Boolean) As Boolean
Declare Function ServerUnshareUsbDev Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long) As Boolean
Declare Function ServerGetUsbDevStatus Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef State As Long, ByRef HostConnect As Variant) As Boolean
Declare Function ServerGetSharedUsbDevNetSettings Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef NetSettings As Variant) As Boolean
Declare Function ServerGetSharedUsbDevIsCrypt Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef Crypt As Boolean) As Boolean
Declare Function ServerGetSharedUsbDevRequiresAuth Lib "RemoteUSB.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef Auth As Boolean) As Boolean

' client
Declare Function ClientAddRemoteDevManually Lib "RemoteUSB.dll" (ByVal NetSettings As Variant) As Boolean
Declare Function ClientAddRemoteDev Lib "RemoteUSB.dll" (ByVal ClientContext As Long, ByVal iIndex As Long) As Boolean
Declare Function ClientStartRemoteDev Lib "RemoteUSB.dll" (ByVal ClientContext As Long, ByVal iIndex As Long, ByVal Reconnect As Boolean, ByVal Password As Variant) As Boolean
Declare Function ClientStopRemoteDev Lib "RemoteUSB.dll" (ByVal ClientContext As Long, ByVal iIndex As Long) As Boolean
Declare Function ClientRemoveRemoteDev Lib "RemoteUSB.dll" (ByVal ClientContext As Long, ByVal iIndex As Long) As Boolean
Declare Function ClientGetStateRemoteDev Lib "RemoteUSB.dll" (ByVal ClientContext As Long, ByVal iIndex As Long, ByRef State As Integer, ByRef RemoteHost As Variant) As Boolean
Declare Function ClientTrafficRemoteDevIsEncrypted Lib "RemoteUSB.dll" (ByVal ClientContext As Long, ByVal iIndex As Long, ByRef Crypt As Boolean) As Boolean
Declare Function ClientRemoteDevRequiresAuth Lib "RemoteUSB.dll" (ByVal ClientContext As Long, ByVal iIndex As Long, ByRef Auth As Boolean) As Boolean

' enum client dev
Declare Function ClientEnumAvailRemoteDevOnServer Lib "RemoteUSB.dll" (ByVal Server As Variant, ByRef FindContext As Long) As Boolean
Declare Function ClientEnumAvailRemoteDev Lib "RemoteUSB.dll" (ByRef FindContext As Long) As Boolean
Declare Function ClientRemoveEnumOfRemoteDev Lib "RemoteUSB.dll" (ByVal FindContext As Long) As Boolean
Declare Function ClientGetRemoteDevNetSettings Lib "RemoteUSB.dll" (ByVal FindContext As Long, ByVal Index As Long, ByRef NetSettings As Variant) As Boolean
Declare Function ClientGetRemoteDevName Lib "RemoteUSB.dll" (ByVal FindContext As Long, ByVal Index As Long, ByRef Name As Variant) As Boolean


