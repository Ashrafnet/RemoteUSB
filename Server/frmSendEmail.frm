VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmSendEmail 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1560
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1560
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
End
Attribute VB_Name = "frmSendEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Send a text email using the MS MAPI control

'The following code sends an email using a MAPI client. You must first reference the "Microsoft MAPI Controls 6.0" and add both the "MAPI Messages" and "MAPI Session" controls to a form.


'Purpose     :  Send an plain text Email
'Inputs      :  sSendTo                         The email address of the recipient of the mail
'               sSubject                        The subject/title of the mail
'               sText                           The mail content
'               [sAttachFile]                   Optional file to attach with the email.
'Outputs     :  Returns True if successful
'Notes       :  First reference then add both the MS MAPI Components to a form
'               In VBA, right click the "Toolbox" dialog then select "addition controls"


Function MailSend(sSendTo As String, sSubject As String, sText As String, Optional sAttachFile As String) As Boolean
    On Error GoTo ErrHandler
    With MAPISession1
        If .SessionID = 0 Then
            .DownLoadMail = False
            .LogonUI = True
            .SignOn
            .NewSession = True
            MAPIMessages1.SessionID = .SessionID
        End If
    End With
    With MAPIMessages1
        .Compose
        .RecipAddress = sSendTo
        .AddressResolveUI = True
        .ResolveName
        .MsgSubject = sSubject
        .MsgNoteText = sText
        If Len(sAttachFile) > 0 And Len(Dir$(sAttachFile)) > 0 Then
            .AttachmentPathName = sAttachFile
        Else
            .AttachmentCount = 0
        End If
        .Send False
    End With
    MailSend = True
    Exit Function
ErrHandler:
    Debug.Print Err.Description
    MailSend = False
    
    
End Function

'Demonstration routine
Sub Test()
    MailSend "webmaster@domain", "MAPI Test", "Test Message!"
    MailSend "admin@domain", "MAPI Test", "Test Message!"
End Sub

