Attribute VB_Name = "u2ec_example"
' enum hub
Declare Function ServerCreateEnumUsbDev Lib "u2ec.dll" (ByRef Context As Long) As Boolean
Declare Function ServerRemoveEnumUsbDev Lib "u2ec.dll" (ByVal Context As Long) As Boolean
Declare Function ServerGetUsbDevFromHub Lib "u2ec.dll" (ByVal Context As Long, ByVal HubContext As Long, ByVal Index As Long, ByRef DevContext As Long) As Boolean
Declare Function ServerUsbDevIsHub Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long) As Boolean
Declare Function ServerUsbDevIsShared Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long) As Boolean
Declare Function ServerUsbDevIsConnected Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long) As Boolean
Declare Function ServerGetUsbDevName Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef Name As Variant) As Boolean

Declare Function SetCallBackOnChangeDevList Lib "u2ec.dll" (ByRef Callback As Long) As Boolean

' server
Declare Function ServerShareUsbDev Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long, ByVal Connectionn As Variant, ByVal Description As Variant, ByVal Auth As Boolean, ByVal Passw As Variant, ByVal Crypt As Boolean) As Boolean
Declare Function ServerUnshareUsbDev Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long) As Boolean
Declare Function ServerGetUsbDevStatus Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef State As Long, ByRef HostConnect As Variant) As Boolean
Declare Function ServerGetSharedUsbDevNetSettings Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef NetSettings As Variant) As Boolean
Declare Function ServerGetSharedUsbDevIsCrypt Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef Crypt As Boolean) As Boolean
Declare Function ServerGetSharedUsbDevRequiresAuth Lib "u2ec.dll" (ByVal Context As Long, ByVal DevContext As Long, ByRef Auth As Boolean) As Boolean

' client
Declare Function ClientAddRemoteDevManually Lib "u2ec.dll" (ByVal NetSettings As Variant) As Boolean
Declare Function ClientAddRemoteDev Lib "u2ec.dll" (ByVal ClientContext As Long, ByVal iIndex As Long) As Boolean
Declare Function ClientStartRemoteDev Lib "u2ec.dll" (ByVal ClientContext As Long, ByVal iIndex As Long, ByVal Reconnect As Boolean, ByVal Password As Variant) As Boolean
Declare Function ClientStopRemoteDev Lib "u2ec.dll" (ByVal ClientContext As Long, ByVal iIndex As Long) As Boolean
Declare Function ClientRemoveRemoteDev Lib "u2ec.dll" (ByVal ClientContext As Long, ByVal iIndex As Long) As Boolean
Declare Function ClientGetStateRemoteDev Lib "u2ec.dll" (ByVal ClientContext As Long, ByVal iIndex As Long, ByRef State As Integer, ByRef RemoteHost As Variant) As Boolean
Declare Function ClientTrafficRemoteDevIsEncrypted Lib "u2ec.dll" (ByVal ClientContext As Long, ByVal iIndex As Long, ByRef Crypt As Boolean) As Boolean
Declare Function ClientRemoteDevRequiresAuth Lib "u2ec.dll" (ByVal ClientContext As Long, ByVal iIndex As Long, ByRef Auth As Boolean) As Boolean

' enum client dev
Declare Function ClientEnumAvailRemoteDevOnServer Lib "u2ec.dll" (ByVal Server As Variant, ByRef FindContext As Long) As Boolean
Declare Function ClientEnumAvailRemoteDev Lib "u2ec.dll" (ByRef FindContext As Long) As Boolean
Declare Function ClientRemoveEnumOfRemoteDev Lib "u2ec.dll" (ByVal FindContext As Long) As Boolean
Declare Function ClientGetRemoteDevNetSettings Lib "u2ec.dll" (ByVal FindContext As Long, ByVal Index As Long, ByRef NetSettings As Variant) As Boolean
Declare Function ClientGetRemoteDevName Lib "u2ec.dll" (ByVal FindContext As Long, ByVal Index As Long, ByRef Name As Variant) As Boolean

