Attribute VB_Name = "MapiVB32"
Option Explicit

'**************************************************************************
'
'
'
' Visual Basic declaration for the MAPI functions.
'
' This file can be loaded into the global module.
'
'
'
'
'**************************************************************************
'

'***************************************************
'   MAPI Message holds information about a message
'***************************************************

'Type MAPIMessage
'    Reserved As Long
'    Subject As String
'    NoteText As String
'    MessageType As String
'    DateReceived As String
'    ConversationID As String
'    flags As Long
'    RecipCount As Long
'    FileCount As Long
'End Type
'Type MAPIMessage
'    Reserved As Long
'    Subject As String
'    NoteText As String
'    MessageType As String
'    DateReceived As String
'    ConversationID As String
'    flFlags As Long
'    lpOriginator As Long
'    nRecipCount As Long
'    lpRecips As Long
'    nFileCount As Long
'    lpFiles As Long
'End Type


'************************************************
'   MAPIRecip holds information about a message
'   originator or recipient
'************************************************

'Type MapiRecip
'    Reserved As Long
'    RecipClass As Long
'    Name As String
'    Address As String
'    EIDSize As Long
'    EntryID As String
'End Type
'Type MapiRecip
'    Reserved   As Long
'    RecipClass As Long
'    Name    As String
'    Address As String
'    EIDSize As Long
'    EntryID As Long 'String
'End Type



'******************************************************
'   MapiFile holds information about file attachments
'******************************************************

'Type MapiFile
'    Reserved As Long
'    flags As Long
'    Position As Long
'    PathName As String
'    FileName As String
'    FileType As String
'End Type
'Type MapiFile
'    Reserved  As Long
'    flFlags   As Long
'    nPosition As Long
'    PathName  As String
'    FileName  As String
'    FileType  As Long 'String
'End Type



'***************************
'   FUNCTION Declarations
'***************************

'ULONG WINAPI MAPISendMail(
'  _In_ LHANDLE       lhSession,
'  _In_ ULONG_PTR     ulUIParam,
'  _In_ lpMapiMessage lpMessage,
'  _In_ FLAGS         flFlags,
'       ULONG ulReserved
');

'Public Declare Function MAPISendMail Lib "mapi32" (ByVal lhSession As Long, ByVal ulUIParam As Long, ByRef lpMessage As MAPIMessage, ByVal Flags As Long, ByVal ulReserved As Long)

'Public Declare Function MAPILogon Lib "MAPI32.DLL" (ByVal UIParam&, ByVal User$, ByVal Password$, ByVal flags&, ByVal Reserved&, Session&) As Long
'Public Declare Function MAPILogoff Lib "MAPI32.DLL" (ByVal Session&, ByVal UIParam&, ByVal flags&, ByVal Reserved&) As Long
'Public Declare Function BMAPIReadMail Lib "MAPI32.DLL" (lMsg&, nRecipients&, nFiles&, ByVal Session&, ByVal UIParam&, MessageID$, ByVal Flag&, ByVal Reserved&) As Long
'Public Declare Function BMAPIGetReadMail Lib "MAPI32.DLL" (ByVal lMsg&, Message As MAPIMessage, Recip() As MapiRecip, File() As MapiFile, Originator As MapiRecip) As Long
'Public Declare Function MAPIFindNext Lib "MAPI32.DLL" Alias "BMAPIFindNext" (ByVal Session&, ByVal UIParam&, MsgType$, SeedMsgID$, ByVal Flag&, ByVal Reserved&, MsgID$) As Long
'Public Declare Function MAPISendDocuments Lib "MAPI32.DLL" (ByVal UIParam&, ByVal DelimStr$, ByVal FilePaths$, ByVal FileNames$, ByVal Reserved&) As Long
'Public Declare Function MAPIDeleteMail Lib "MAPI32.DLL" (ByVal Session&, ByVal UIParam&, ByVal MsgID$, ByVal flags&, ByVal Reserved&) As Long
'Public Declare Function MAPISendMail Lib "MAPI32.DLL" Alias "BMAPISendMail" (ByVal Session&, ByVal UIParam&, Message As MAPIMessage, Recipient() As MapiRecip, File() As MapiFile, ByVal flags&, ByVal Reserved&) As Long
'Public Declare Function MAPISaveMail Lib "MAPI32.DLL" Alias "BMAPISaveMail" (ByVal Session&, ByVal UIParam&, Message As MAPIMessage, Recipient() As MapiRecip, File() As MapiFile, ByVal flags&, ByVal Reserved&, MsgID$) As Long
'Public Declare Function BMAPIAddress Lib "MAPI32.DLL" (lInfo&, ByVal Session&, ByVal UIParam&, Caption$, ByVal nEditFields&, Label$, nRecipients&, Recip() As MapiRecip, ByVal flags&, ByVal Reserved&) As Long
'Public Declare Function BMAPIGetAddress Lib "MAPI32.DLL" (ByVal lInfo&, ByVal nRecipients&, recipients() As MapiRecip) As Long
'Public Declare Function MAPIDetails Lib "MAPI32.DLL" Alias "BMAPIDetails" (ByVal Session&, ByVal UIParam&, Recipient As MapiRecip, ByVal flags&, ByVal Reserved&) As Long
'Public Declare Function MAPIResolveName Lib "MAPI32.DLL" Alias "BMAPIResolveName" (ByVal Session&, ByVal UIParam&, ByVal UserName$, ByVal flags&, ByVal Reserved&, Recipient As MapiRecip) As Long



'**************************
'   CONSTANT Declarations
'**************************
'

'Public Const ERROR_SUCCESS                   As Long = 0
'Public Const MAPI_USER_ABORT                 As Long = 1
'Public Const MAPI_E_USER_ABORT               As Long = MAPI_USER_ABORT
'Public Const MAPI_E_FAILURE                  As Long = 2
'Public Const MAPI_E_LOGIN_FAILURE            As Long = 3
'Public Const MAPI_E_LOGON_FAILURE            As Long = MAPI_E_LOGIN_FAILURE
'Public Const MAPI_E_DISK_FULL                As Long = 4
'Public Const MAPI_E_INSUFFICIENT_MEMORY      As Long = 5
'Public Const MAPI_E_BLK_TOO_SMALL            As Long = 6
'Public Const MAPI_E_TOO_MANY_SESSIONS        As Long = 8
'Public Const MAPI_E_TOO_MANY_FILES           As Long = 9
'Public Const MAPI_E_TOO_MANY_RECIPIENTS      As Long = 10
'Public Const MAPI_E_ATTACHMENT_NOT_FOUND     As Long = 11
'Public Const MAPI_E_ATTACHMENT_OPEN_FAILURE  As Long = 12
'Public Const MAPI_E_ATTACHMENT_WRITE_FAILURE As Long = 13
'Public Const MAPI_E_UNKNOWN_RECIPIENT        As Long = 14
'Public Const MAPI_E_BAD_RECIPTYPE            As Long = 15
'Public Const MAPI_E_NO_MESSAGES              As Long = 16
'Public Const MAPI_E_INVALID_MESSAGE          As Long = 17
'Public Const MAPI_E_TEXT_TOO_LARGE           As Long = 18
'Public Const MAPI_E_INVALID_SESSION          As Long = 19
'Public Const MAPI_E_TYPE_NOT_SUPPORTED       As Long = 20
'Public Const MAPI_E_AMBIGUOUS_RECIPIENT      As Long = 21
'Public Const MAPI_E_AMBIG_RECIP              As Long = MAPI_E_AMBIGUOUS_RECIPIENT
'Public Const MAPI_E_MESSAGE_IN_USE           As Long = 22
'Public Const MAPI_E_NETWORK_FAILURE          As Long = 23
'Public Const MAPI_E_INVALID_EDITFIELDS       As Long = 24
'Public Const MAPI_E_INVALID_RECIPS           As Long = 25
'Public Const MAPI_E_NOT_SUPPORTED            As Long = 26
'
'Public Const MAPI_ORIG As Long = 0
'Public Const MAPI_TO   As Long = 1
'Public Const MAPI_CC   As Long = 2
'Public Const MAPI_BCC  As Long = 3


'***********************
'   FLAG Declarations
'***********************

'* MAPILogon() flags *
'
'Global Const MAPI_LOGON_UI = &H1
'Global Const MAPI_NEW_SESSION = &H2
'Global Const MAPI_FORCE_DOWNLOAD = &H1000
'
''* MAPILogoff() flags *
'
'Global Const MAPI_LOGOFF_SHARED = &H1
'Global Const MAPI_LOGOFF_UI = &H2
'
''* MAPISendMail() flags *
'
'Global Const MAPI_DIALOG = &H8
'
''* MAPIFindNext() flags *
'
'Global Const MAPI_UNREAD_ONLY = &H20
'Global Const MAPI_GUARANTEE_FIFO = &H100
'
''* MAPIReadMail() flags *
'
'Global Const MAPI_ENVELOPE_ONLY = &H40
'Global Const MAPI_PEEK = &H80
'Global Const MAPI_BODY_AS_FILE = &H200
'Global Const MAPI_SUPPRESS_ATTACH = &H800
'
''* MAPIDetails() flags *
'
'Global Const MAPI_AB_NOMODIFY = &H400
'
''* Attachment flags *
'
'Global Const MAPI_OLE = &H1
'Global Const MAPI_OLE_STATIC = &H2
'
''* MapiMessage flags *
'
'Global Const MAPI_UNREAD = &H1
'Global Const MAPI_RECEIPT_REQUESTED = &H2
'Global Const MAPI_SENT = &H4

'Public Function CopyFiles(MfIn As MapiFile, MfOut As MapiFile) As Long
'
'    MfOut.FileName = MfIn.FileName
'    MfOut.PathName = MfIn.PathName
'    MfOut.Reserved = MfIn.Reserved
'    MfOut.Flags = MfIn.Flags
'    MfOut.position = MfIn.position
'    MfOut.FileType = MfIn.FileType
'    CopyFiles = 1&
'
'End Function
'
'Public Function CopyRecipient(MrIn As MapiRecip, MrOut As MapiRecip) As Long
'
'    MrOut.Name = MrIn.Name
'    MrOut.Address = MrIn.Address
'    MrOut.EIDSize = MrIn.EIDSize
'    MrOut.EntryID = MrIn.EntryID
'    MrOut.Reserved = MrIn.Reserved
'    MrOut.RecipClass = MrIn.RecipClass
'
'    CopyRecipient = 1&
'
'End Function
'
'Public Function MAPIAddress(Session As Long, UIParam As Long, Caption As String, _
'nEditFields As Long, Label As String, nRecipients As Long, Recips() As _
'MapiRecip, Flags As Long, Reserved As Long) As Long
'
'
'    Dim Info&
'    Dim rc&
'    Dim nRecips As Long
'
'    ReDim Rec(0 To nRecipients) As MapiRecip
'    ' Use local variable since BMAPIAddress changes the passed value
'    nRecips = nRecipients
'
'    '*****************************************************
'    ' Copy input recipient structure into local
'    ' recipient structure used as input to BMAPIAddress
'    '*****************************************************
'
'    For i = 0 To nRecipients - 1
'        Ignore& = CopyRecipient(Recips(i), Rec(i))
'    Next i
'
'    rc& = BMAPIAddress(Info&, Session&, UIParam&, Caption$, nEditFields&, Label$, nRecips&, Rec(), Flags, 0&)
'
'    If (rc& = SUCCESS_SUCCESS) Then
'
'        '**************************************************
'        ' New recipients are now in the memory referenced
'        ' by Info (HANDLE). nRecipients is the number of
'        ' new recipients.
'        '**************************************************
'        nRecipients = nRecips     ' Copy back to parameter
'
'        If (nRecipients > 0) Then
'            ReDim Rec(0 To nRecipients - 1) As MapiRecip
'            rc& = BMAPIGetAddress(Info&, nRecipients&, Rec())
'
'            '*********************************************
'            ' Copy local recipient structure to
'            ' recipient structure passed as procedure
'            ' parameter. This is necessary because
'            ' VB doesn't seem to work properly when
'            ' the procedure parameter gets passed
'            ' directory to the BMAPI.DLL Address routine
'            '*********************************************
'
'            ReDim Recips(0 To nRecipients - 1) As MapiRecip
'
'            For i = 0 To nRecipients - 1
'                Ignore& = CopyRecipient(Rec(i), Recips(i))
'            Next i
'
'        End If
'
'    End If
'
'    MAPIAddress = rc&
'
'End Function
'
'Public Function MAPIReadMail(Session As Long, UIParam As Long, MessageID As _
'String, Flags As Long, Reserved As Long, Message As MAPIMessage, Orig As _
'MapiRecip, RecipsOut() As MapiRecip, FilesOut() As MapiFile) As Long
'
'    Dim Info&
'    Dim nFiles&, nRecips&
'
'    rc& = BMAPIReadMail(Info&, nRecips, nFiles, Session, 0, MessageID, _
'Flags, Reserved)
'
'    If (rc& = SUCCESS_SUCCESS) Then
'
'        'Message is now read into the handles array. We have to redim the
'        'arrays and read the information in.
'
'        If (nRecips = 0) Then nRecips = 1
'        If (nFiles = 0) Then nFiles = 1
'
'        ReDim Recips(0 To nRecips - 1) As MapiRecip
'        ReDim Files(0 To nFiles - 1) As MapiFile
'
'        rc& = BMAPIGetReadMail(Info&, Message, Recips(), Files(), Orig)
'
'        '*******************************************
'        ' Copy Recipient and File structures from
'        ' Local structures to those passed as
'        ' parameters
'        '*******************************************
'
'        ReDim FilesOut(0 To nFiles - 1) As MapiFile
'        ReDim RecipsOut(0 To nRecips - 1) As MapiRecip
'
'        For i = 0 To nRecips - 1
'            Ignore& = CopyRecipient(Recips(i), RecipsOut(i))
'        Next i
'
'        For i = 0 To nFiles - 1
'            Ignore& = CopyFiles(Files(i), FilesOut(i))
'        Next i
'
'    End If
'
'    MAPIReadMail = rc&
'
'End Function
'
'Public Sub SendMail()
'    Dim oMsg As MAPIMessage
'    Dim oRecipients(1) As MapiRecip
'    Dim oAttachments(0) As MapiFile
'    Dim lSession As Long
'    Dim lResult As Long
'
'    ' Logon
'    lResult = MAPILogon(0, "", "", MAPI_LOGON_UI + MAPI_NEW_SESSION, 0, lSession)
'
'    If lResult <> 0 Then
'        MsgBox "Logon failed. Result = " & lResult
'        Exit Sub
'    End If
'
'    ' Fill out the message
'    With oMsg
'        .Reserved = 0
'        .NoteText = "Test message body"
'        .FileCount = 0  'no attachments
'        .RecipCount = 1  'only 1 recipient
'        .Subject = "Test message"
'    End With
'
'    ' Fill out the recipient
'    With oRecipients(0)
'        ' TODO: Change "test@online.microsoft.com" to the address you want to send to
'        .Name = "test@online.microsoft.com"
'        .RecipClass = MAPI_TO
'        .Reserved = 0
'    End With
'
'    lResult = MAPIResolveName(lSession, 0, oRecipients(0).Name, 0, 0, oRecipients(0))
'
'    If lResult <> 0 Then
'        MsgBox "MAPIResolveName failed. Result = " & lResult
'        Exit Sub
'    End If
'
'    ' Send the message
'    lResult = MAPISendMail(0, 0, oMsg, oRecipients, oAttachments, 0, 0&)
'
'    If lResult = 0 Then
'        MsgBox ("Message sent!")
'    Else
'        MsgBox "Message not sent! Result = " & lResult
'    End If
'
'    ' Log off
'    lResult = MAPILogoff(lSession, 0, 0, 0)
'
'    If lResult <> 0 Then
'        MsgBox "Logoff failed. Result = " & lResult
'    End If
'
'End Sub
