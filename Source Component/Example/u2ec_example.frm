VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "U2EC example"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "TCP and callback settings"
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   5895
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   $"u2ec_example.frx":0000
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.TextBox Desc 
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add remote device manually (in case client PC cannot see server)"
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   8280
      Width           =   6975
      Begin VB.TextBox ClientAddManualNetSettings 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   5175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add client manual"
         Height          =   735
         Left            =   5520
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   $"u2ec_example.frx":00ED
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find shared devices on remote server"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   5295
      Begin VB.CommandButton AddClient 
         Caption         =   "Add client"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox RemoteServer 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Text            =   "localhost"
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Find 
         Caption         =   "Find"
         Height          =   315
         Left            =   3960
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Note: enter server name and press ""Find"". Choose one device from list and click ""Add device"""
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Remote server:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton FindLocal 
      Caption         =   "Show added remote devices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton StatusClient 
      Caption         =   "Get status"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Disconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Connect 
      Caption         =   "Connect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton RemoveClient 
      Caption         =   "Remove device"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   7800
      Width           =   1335
   End
   Begin VB.ListBox ListClient 
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   6975
   End
   Begin VB.CommandButton State 
      Caption         =   "Get status"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Unshare 
      Caption         =   "Unshare"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Share 
      Caption         =   "Share"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6975
   End
   Begin VB.Label Label8 
      Caption         =   "OR"
      Height          =   255
      Left            =   5520
      TabIndex        =   29
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Центровка
      BackColor       =   &H80000002&
      Caption         =   "Share local USB devices  (server connection)"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label6 
      Caption         =   "Additional device description:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Центровка
      BackColor       =   &H80000002&
      Caption         =   "Connect to remote shared USB devices (client connection)"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   4920
      Width           =   7215
   End
   Begin VB.Label Label4 
      Caption         =   "USB devices list:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Handle As Long
Dim HandleDevs(256) As Long

Dim HandleClient As Long
Dim HandleDevsClient(256) As Long

Dim GlobalIndex As Long

Public Sub EnumHubToList(Context As Long, HubContext As Long, Level As Integer)
    Dim HandleDev As Long
    Dim Index As Long
    Dim strName As Variant
    Dim strNameDev As Variant
    Dim Connect As Variant
    Dim State As Long
    Index = 0
    
    If Context <> HubContext Then
        ret = ServerGetUsbDevName(Context, HubContext, strName)
        Dim Space As String
        Space = String(Level, " ")
        List1.AddItem (Space + strName)
        HandleDevs(GlobalIndex) = HubContext
        GlobalIndex = GlobalIndex + 1
    End If
    
    Level = Level + 1
    While ServerGetUsbDevFromHub(Context, HubContext, Index, HandleDev)
        If ServerUsbDevIsHub(Context, HandleDev) Then
            EnumHubToList Context, HandleDev, Level
        Else
            ret = ServerGetUsbDevName(Context, HandleDev, strName)
            Space = String(Level, " ")
            
            strNameDev = Space + strName
            Connect = ""
            If ServerGetUsbDevStatus(Context, HandleDev, State, Connect) Then
                If ServerGetSharedUsbDevNetSettings(Context, HandleDev, strName) Then
                    strNameDev = strNameDev + " / Shared - " + strName
                    If (Len(Connect) > 0) And (State = 2) Then
                        strNameDev = strNameDev + " / connected to " + Connect
                    End If
                End If
            End If
            
            List1.AddItem (strNameDev)
            
            HandleDevs(GlobalIndex) = HandleDev
            GlobalIndex = GlobalIndex + 1
        End If
        Index = Index + 1
    Wend
    Level = Level - 1

End Sub

Private Sub AddClient_Click()
    Dim Name As Variant
    Dim Index As Integer
    Index = ListClient.ListIndex
    If ClientAddRemoteDev(HandleClient, Index) Then
        If BuildName(HandleClient, Index, Name) Then
            ListClient.List(Index) = Name
        End If
    End If
       
    ListClient_Click
End Sub

Private Sub Command1_Click()
    ServerRemoveEnumUsbDev (Handle)
    List1.Clear
    GlobalIndex = 0

    If ServerCreateEnumUsbDev(Handle) Then
        EnumHubToList Handle, Handle, 0
    End If
    
    List1_Click
End Sub

Private Sub Command2_Click()
    ClientAddRemoteDevManually (ClientAddManualNetSettings.Text)
    
    ListClient_Click
End Sub


Private Sub Connect_Click()
    ret = ClientStartRemoteDev(HandleClient, ListClient.ListIndex, True, "")
    ListClient_Click
End Sub

Private Sub Disconnect_Click()
    ret = ClientStopRemoteDev(HandleClient, ListClient.ListIndex)
    ListClient_Click
End Sub

Private Sub Find_Click()
    Dim Index As Long
    Dim Name As Variant
    
    ListClient.Clear
    
    If HandleClient <> 0 Then
        ClientRemoveEnumOfRemoteDev (HandleClient)
    End If
    
    If ClientEnumAvailRemoteDevOnServer(RemoteServer.Text, HandleClient) Then
        Index = 0
        While BuildName(HandleClient, Index, Name)
            ListClient.AddItem (Name)
            Index = Index + 1
        Wend
    End If
    
    ListClient_Click
End Sub

Private Function BuildName(ByVal ClientContext As Long, ByVal iIndex As Long, ByRef Name As Variant) As Boolean
    Dim NetSettings As Variant
    Dim Host As Variant
    Dim State As Integer
    bRet = False
    
    If ClientGetRemoteDevName(ClientContext, iIndex, Name) Then
        BuildName = True
        Name = IIf(Len(Name) = 0, "Unknown", Name)
        If ClientGetStateRemoteDev(ClientContext, iIndex, State, Host) Then
            If ClientGetRemoteDevNetSettings(ClientContext, iIndex, NetSettings) Then
                Name = Name + " / " + NetSettings
            End If
        End If
    Else
        BuildName = False
    End If
End Function

Private Sub FindLocal_Click()
    Dim Index As Long
    Dim Name As Variant
    Dim NetSettings As Variant
    Dim State As Long
    Dim RemoteHost As Variant
    
    ListClient.Clear
    
    If HandleClient <> 0 Then
        ClientRemoveEnumOfRemoteDev (HandleClient)
    End If
    
    If ClientEnumAvailRemoteDev(HandleClient) Then
        Index = 0
        While BuildName(HandleClient, Index, Name)
            ListClient.AddItem (Name)
            Index = Index + 1
        Wend
    End If
    
    ListClient_Click
End Sub

Private Sub Form_Load()
    Handle = 0
    GlobalIndex = 0
    HandleClient = 0
    If ServerCreateEnumUsbDev(Handle) Then
        EnumHubToList Handle, Handle, 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ServerRemoveEnumUsbDev (Handle)
    If HandleClient <> 0 Then
        ClientRemoveEnumOfRemoteDev (HandleClient)
    End If
End Sub


Private Sub List1_Click()
    ret = False
    If List1.ListIndex <> -1 Then
        ret = ServerUsbDevIsHub(Handle, HandleDevs(List1.ListIndex))
    Else
        Share.Enabled = False
        Unshare.Enabled = False
        Label1.Enabled = False
        Text1.Enabled = False
        State.Enabled = False
        Exit Sub
    End If
    
    If Not ret Then
        ret = ServerUsbDevIsShared(Handle, HandleDevs(List1.ListIndex))
        Share.Enabled = Not ret
        Unshare.Enabled = ret
        Label1.Enabled = Not ret
        Text1.Enabled = Not ret
        State.Enabled = ret
    Else
        Share.Enabled = False
        Unshare.Enabled = False
        Label1.Enabled = False
        Text1.Enabled = False
        State.Enabled = False
    End If
End Sub

Private Sub ListClient_Click()
    Dim State As Integer
    Dim ret As Boolean

    ret = (ListClient.ListIndex <> -1)
    If ret Then
        ret = ClientGetStateRemoteDev(ByVal HandleClient, ListClient.ListIndex, State, Remote)
    End If
    
    AddClient.Enabled = Not ret
    RemoveClient.Enabled = ret
    Connect.Enabled = ret And State = 0
    Disconnect.Enabled = ret And State > 0
    StatusClient.Enabled = ret

End Sub

Private Sub ListClient_LostFocus()
    ListClient_Click
End Sub

Private Sub RemoveClient_Click()
    Dim Name As Variant
    Dim Index As Integer
    Index = ListClient.ListIndex
    If ClientRemoveRemoteDev(HandleClient, Index) Then
        If BuildName(HandleClient, Index, Name) Then
            ListClient.List(Index) = Name
        End If
    End If
    'ret = ClientRemoveRemoteDev(HandleClient, ListClient.ListIndex)
    ListClient_Click
End Sub

Private Sub Share_Click()
    ret = ServerShareUsbDev(Handle, HandleDevs(List1.ListIndex), Text1.Text, Desc.Text, False, "", False)
    List1_Click
End Sub

Private Sub State_Click()
    Dim State As Long
    Dim Host As Variant
    Dim NetSettins As Variant
    Dim result As String
        
    ret = ServerGetUsbDevStatus(Handle, HandleDevs(List1.ListIndex), State, Host)
    ret = ServerGetSharedUsbDevNetSettings(Handle, HandleDevs(List1.ListIndex), NetSettins)
    
    result = IIf(State = 2, "connected", "waiting for connection") + "/" + NetSettins + IIf(State, "/" + Host, "")
    MsgBox (result)
End Sub


Private Sub StatusClient_Click()
    Dim State As Integer
    Dim Host As Variant
    Dim NetSettins As Variant
    Dim result As String
    
    If ClientGetStateRemoteDev(HandleClient, ListClient.ListIndex, State, Host) Then
        ret = ClientGetRemoteDevNetSettings(HandleClient, ListClient.ListIndex, NetSettins)
    
        result = "added"
        If State = 1 Then
            result = "connecting"
        ElseIf State = 2 Then
            result = "connected"
        End If
        
        result = result + " / " + NetSettins
        result = result + IIf(State = 2, "/" + Host, "")
        
        MsgBox (result)
    End If
End Sub

Private Sub Unshare_Click()
    ret = ServerUnshareUsbDev(Handle, HandleDevs(List1.ListIndex))
    List1_Click
End Sub
