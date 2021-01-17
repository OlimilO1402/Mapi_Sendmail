VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StdMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ERROR_SUCCESS                   As Long = 0
Private Const MAPI_E_USER_ABORT               As Long = 1
Private Const MAPI_E_FAILURE                  As Long = 2
Private Const MAPI_E_LOGIN_FAILURE            As Long = 3
Private Const MAPI_E_DISK_FULL                As Long = 4
Private Const MAPI_E_INSUFFICIENT_MEMORY      As Long = 5
Private Const MAPI_E_BLK_TOO_SMALL            As Long = 6
                                             'As Long = 7 ?
Private Const MAPI_E_TOO_MANY_SESSIONS        As Long = 8
Private Const MAPI_E_TOO_MANY_FILES           As Long = 9
Private Const MAPI_E_TOO_MANY_RECIPIENTS      As Long = 10
Private Const MAPI_E_ATTACHMENT_NOT_FOUND     As Long = 11
Private Const MAPI_E_ATTACHMENT_OPEN_FAILURE  As Long = 12
Private Const MAPI_E_ATTACHMENT_WRITE_FAILURE As Long = 13
Private Const MAPI_E_UNKNOWN_RECIPIENT        As Long = 14
Private Const MAPI_E_BAD_RECIPTYPE            As Long = 15
Private Const MAPI_E_NO_MESSAGES              As Long = 16
Private Const MAPI_E_INVALID_MESSAGE          As Long = 17
Private Const MAPI_E_TEXT_TOO_LARGE           As Long = 18
Private Const MAPI_E_INVALID_SESSION          As Long = 19
Private Const MAPI_E_TYPE_NOT_SUPPORTED       As Long = 20
Private Const MAPI_E_AMBIGUOUS_RECIPIENT      As Long = 21
Private Const MAPI_E_MESSAGE_IN_USE           As Long = 22
Private Const MAPI_E_NETWORK_FAILURE          As Long = 23
Private Const MAPI_E_INVALID_EDITFIELDS       As Long = 24
Private Const MAPI_E_INVALID_RECIPS           As Long = 25
Private Const MAPI_E_NOT_SUPPORTED            As Long = 26

Private Const MAPI_LOGON_UI        As Long = &H1
Private Const MAPI_NEW_SESSION     As Long = &H2
Private Const MAPI_FORCE_DOWNLOAD  As Long = &H1000

'* MAPILogoff flags *
Private Const MAPI_LOGOFF_SHARED   As Long = &H1
Private Const MAPI_LOGOFF_UI       As Long = &H2

'* MAPISendMail flags *
Private Const MAPI_DIALOG          As Long = &H8

'* MAPIFindNext flags *
Private Const MAPI_UNREAD_ONLY     As Long = &H20
Private Const MAPI_GUARANTEE_FIFO  As Long = &H100

'* MAPIReadMail flags *
Private Const MAPI_ENVELOPE_ONLY   As Long = &H40
Private Const MAPI_PEEK            As Long = &H80
Private Const MAPI_BODY_AS_FILE    As Long = &H200
Private Const MAPI_SUPPRESS_ATTACH As Long = &H800

'* MAPIDetails flags *
Private Const MAPI_AB_NOMODIFY     As Long = &H400

'* Attachment flags *
Private Const MAPI_OLE             As Long = &H1
Private Const MAPI_OLE_STATIC      As Long = &H2

'* MapiMessage flags *
Private Const MAPI_UNREAD             As Long = &H1
Private Const MAPI_RECEIPT_REQUESTED  As Long = &H2
Private Const MAPI_SENT               As Long = &H4

'* MAPISendMailW flags *
Private Const MAPI_FORCE_UNICODE     As Long = &H40000

Private Const MAPI_ORIG As Long = 0
Private Const MAPI_TO   As Long = 1
Private Const MAPI_CC   As Long = 2
Private Const MAPI_BCC  As Long = 3
'typedef struct {
'  ULONG            ulReserved;
'  PWSTR            lpszSubject;
'  PWSTR            lpszNoteText;
'  PWSTR            lpszMessageType;
'  PWSTR            lpszDateReceived;
'  PWSTR            lpszConversationID;
'  FLAGS            flFlags;
'  lpMapiRecipDescW lpOriginator;
'  ULONG            nRecipCount;
'  lpMapiRecipDescW lpRecips;
'  ULONG            nFileCount;
'  lpMapiFileDescW  lpFiles;
'} MapiMessageW, *lpMapiMessageW;
'Private Type MapiMessageW
'    ulReserved         As Long
'    lpszSubject        As String
'    lpszNoteText       As String
'    lpszMessageType    As String
'    lpszDateReceived   As String
'    lpszConversationID As String
'    flFlags            As Long
'    lpOriginator       As lpMapiRecipDescW
'    nRecipCount        As Long
'    lpRecips           As lpMapiRecipDescW
'    nFileCount         As Long
'    lpFiles            As lpMapiFileDescW
'End Type
Private Type MAPIMessage 'A and W
    Reserved       As Long
    Subject        As String
    NoteText       As String
    MessageType    As String
    DateReceived   As String
    ConversationID As String
    flFlags        As Long
    lpOriginator   As Long
    nRecipCount    As Long
    lpRecips       As Long
    nFileCount     As Long
    lpFiles        As Long
End Type
Private Type MapiRecip 'Desc 'A and W
    Reserved   As Long
    RecipClass As Long
    Name       As String
    Address    As String
    EIDSize    As Long
    EntryID    As Long
End Type
Private Type MapiFile 'Desc 'A and W
    Reserved  As Long
    flFlags   As Long
    nPosition As Long
    PathName  As String
    FileName  As String
    FileType  As Long 'Pointer to a MapiFileTagExt-structure, will not be used here
End Type
'Private Type MapiFileTagExt
'    ulReserved As Long
'    cbTag As Long
'    lpTag As String 'LPBYTE
'    cbEncoding As Long
'    lpEncoding As String 'LPBYTE
'End Type
Private m_HMapiS  As Long
Private m_Recps() As MapiRecip
Private m_Files() As MapiFile
Private m_Email   As MAPIMessage
Private m_IsUnicode As Boolean
Private Declare Function MAPISendMail Lib "mapi32" (ByVal lhSession As Long, ByVal ulUIParam As Long, ByVal lpMessage As Long, ByVal flags As Long, ByVal ulReserved As Long) As Long
Private Declare Function MAPISendMailW Lib "mapi32" (ByVal lhSession As Long, ByVal ulUIParam As Long, ByVal lpMessage As Long, ByVal flags As Long, ByVal ulReserved As Long) As Long
Private Declare Function MAPILogon Lib "mapi32" (ByVal UIParam As Long, ByVal User As String, ByVal password As String, ByVal flags As Long, ByVal Reserved As Long, ByRef Session_out As Long) As Long
'

Private Sub Class_Initialize()
    ReDim m_Recps(0)
    ReDim m_Files(0)
End Sub
Public Sub Login(UserName As String, password As String)
    Dim flags As Long: flags = MAPI_LOGON_UI Or MAPI_NEW_SESSION
    Dim hr    As Long:    hr = MAPILogon(0, UserName, password, flags, 0, m_HMapiS)
    If hr <> ERROR_SUCCESS Then ShowError hr, "Login", UserName & ": " & password
End Sub
Public Property Get IsUnicode() As Boolean
    IsUnicode = m_IsUnicode
End Property
Public Property Let IsUnicode(value As Boolean)
    m_IsUnicode = value
End Property
Public Property Get Recipient(Optional ByVal Index As Long = 0) As String
    With m_Recps(Index)
        Recipient = IIf(m_IsUnicode, .Name & "<" & .Address & ">", StrConv(.Name, vbUnicode) & "<" & StrConv(.Address, vbUnicode) & ">")
'        If m_IsUnicode Then
'            Recipient = .Name & "<" & .Address & ">"
'        Else
'            Recipient = StrConv(.Name, vbUnicode) & "<" & StrConv(.Address, vbUnicode) & ">"
'        End If
    End With
End Property
Public Property Let RecipientName(ByVal Index As Long, value As String)
    m_Recps(Index).Name = value
End Property
Public Property Let RecipientEmailaddress(ByVal Index As Long, value As String)
    m_Recps(Index).Address = value
End Property
Public Sub RecipientAddTo(aEmailAddress As String, Optional aName As String = "")
    PrivateRecipientAdd MAPI_TO, aEmailAddress, aName
End Sub
Public Sub RecipientAddCC(aEmailAddress As String, Optional aName As String = "")
    PrivateRecipientAdd MAPI_CC, aEmailAddress, aName
End Sub
Public Sub RecipientAddBCC(aEmailAddress As String, Optional aName As String = "")
    PrivateRecipientAdd MAPI_BCC, aEmailAddress, aName
End Sub

Private Sub PrivateRecipientAdd(ByVal rClass As Long, aEmailAddress As String, Optional aName As String = "")
    ReDim Preserve m_Recps(m_Email.nRecipCount)
    With m_Recps(m_Email.nRecipCount)
        'If m_IsUnicode Then
        .Name = IIf(m_IsUnicode, aName, StrConv(aName, vbFromUnicode))
        '.Name = StrConv(aName, vbFromUnicode)
        '.Name = aName
        '.Address = StrConv(aEmailAddress, vbFromUnicode)
        .Address = IIf(m_IsUnicode, aEmailAddress, StrConv(aEmailAddress, vbFromUnicode))
        '.Address = aEmailAddress
        .RecipClass = rClass
    End With
    m_Email.nRecipCount = m_Email.nRecipCount + 1
    m_Email.lpRecips = VarPtr(m_Recps(0))
End Sub
Public Property Get RecipientCount() As Long
    RecipientCount = m_Email.nRecipCount
End Property

Public Sub FileAdd(aPFN As String)
    ReDim Preserve m_Files(m_Email.nFileCount)
    With m_Files(m_Email.nFileCount)
        .PathName = IIf(m_IsUnicode, aPFN, StrConv(aPFN, vbFromUnicode))
        '.PathName = StrConv(aPFN, vbFromUnicode)
        .nPosition = &HFFFFFFFF
    End With
    m_Email.nFileCount = m_Email.nFileCount + 1
    m_Email.lpFiles = VarPtr(m_Files(0))
End Sub
Public Property Get FileCount() As Long
    FileCount = m_Email.nFileCount
End Property

Public Property Get Subject() As String
    Subject = IIf(m_IsUnicode, m_Email.Subject, StrConv(m_Email.Subject, vbUnicode))
    'Subject = StrConv(m_Email.Subject, vbUnicode)
End Property
Public Property Let Subject(value As String)
    m_Email.Subject = IIf(m_IsUnicode, value, StrConv(value, vbFromUnicode))
    'm_Email.Subject = StrConv(value, vbFromUnicode)
End Property

Public Property Get BodyText() As String
    BodyText = IIf(m_IsUnicode, m_Email.NoteText, StrConv(m_Email.NoteText, vbUnicode))
    'BodyText = StrConv(m_Email.NoteText, vbUnicode)
End Property
Public Property Let BodyText(value As String)
    m_Email.NoteText = IIf(m_IsUnicode, value, StrConv(value, vbFromUnicode))
    'm_Email.NoteText = StrConv(value, vbFromUnicode)
End Property

Public Sub Start()
    Dim flags As Long: flags = MAPI_DIALOG
    Dim hr    As Long
    If m_IsUnicode Then
        hr = MAPISendMailW(m_HMapiS, 0, VarPtr(m_Email), flags, 0)
    Else
        hr = MAPISendMail(m_HMapiS, 0, VarPtr(m_Email), flags, 0)
    End If
    If hr <> ERROR_SUCCESS Then ShowError hr, "Start"
End Sub

Friend Function ToStr() As String
    ToStr = Me.Recipient & vbCrLf & Me.Subject & vbCrLf & Me.BodyText
End Function

Private Function ShowError(hr As Long, fncname As String, Optional addInfo As String) As VbMsgBoxResult
    If hr = 0 Then Exit Function
    ShowError = MsgBox(GetMAPIErrorMessage(hr, m_Email) & vbCrLf & addInfo, vbInformation Or vbOKCancel, "MAPI-Fehler")
End Function
Private Function GetMAPIErrorMessage(ByVal hr As Long, eml As MAPIMessage) As String
    Dim s As String
    Select Case hr
    Case MAPI_E_USER_ABORT:               s = "User aborted"
    Case MAPI_E_FAILURE:                  s = "Failure"
    Case MAPI_E_LOGIN_FAILURE:            s = "Login failure"
    Case MAPI_E_DISK_FULL:                s = "Disk full"
    Case MAPI_E_INSUFFICIENT_MEMORY:      s = "Insufficient memory"
    Case MAPI_E_BLK_TOO_SMALL:            s = "Bulk too small"
    Case MAPI_E_TOO_MANY_SESSIONS:        s = "Too many sessions" & vbCrLf & m_HMapiS
    Case MAPI_E_TOO_MANY_FILES:           s = "Too many files: " & eml.nFileCount
    Case MAPI_E_TOO_MANY_RECIPIENTS:      s = "Too many recipients: " & eml.nRecipCount
    Case MAPI_E_ATTACHMENT_NOT_FOUND:     s = "Attachment not found: " & vbCrLf & GetAttachments
    Case MAPI_E_ATTACHMENT_OPEN_FAILURE:  s = "Attachment open failure: " & vbCrLf & GetAttachments
    Case MAPI_E_ATTACHMENT_WRITE_FAILURE: s = "Attachment write failure: " & vbCrLf & GetAttachments
    Case MAPI_E_UNKNOWN_RECIPIENT:        s = "Unkown recipient: " & vbCrLf & GetRecipients
    Case MAPI_E_BAD_RECIPTYPE:            s = "Bad recipient type: " & vbCrLf & GetRecipients
    Case MAPI_E_NO_MESSAGES:              s = "No messages: " & vbCrLf & eml.NoteText
    Case MAPI_E_INVALID_MESSAGE:          s = "Invalid messsage: " & vbCrLf & eml.NoteText
    Case MAPI_E_TEXT_TOO_LARGE:           s = "Text too large" & vbCrLf & LenB(eml.NoteText)
    Case MAPI_E_INVALID_SESSION:          s = "Invalid session: " & m_HMapiS
    Case MAPI_E_TYPE_NOT_SUPPORTED:       s = "Type not supported"
    Case MAPI_E_AMBIGUOUS_RECIPIENT:      s = "Ambiguous recipient: " & vbCrLf & GetRecipients
    Case MAPI_E_MESSAGE_IN_USE:           s = "Message in use"
    Case MAPI_E_NETWORK_FAILURE:          s = "NetWork failure"
    Case MAPI_E_INVALID_EDITFIELDS:       s = "Invalid Edit Fields"
    Case MAPI_E_INVALID_RECIPS:           s = "Invalid Recipients: " & vbCrLf & GetRecipients
    Case MAPI_E_NOT_SUPPORTED:            s = "Not supported"
    End Select
    GetMAPIErrorMessage = s
End Function
Private Function GetAttachments() As String
    Dim s As String
    Dim i As Long
    For i = 0 To UBound(m_Files)
        With m_Files(i)
            If m_IsUnicode Then
                s = s & .PathName & " " & .FileName & vbCrLf
            Else
                s = s & StrConv(.PathName, vbUnicode) & " " & StrConv(.FileName, vbUnicode) & vbCrLf
            End If
        End With
        If i = 1000 Then Exit For
    Next
    GetAttachments = s
End Function
Private Function GetRecipients() As String
    Dim sTo As String, sCC As String, sBC As String, sBad As String
    Dim i As Long
    For i = 0 To UBound(m_Recps)
        Select Case m_Recps(i).RecipClass
        Case MAPI_TO:  sTo = sTo & Me.Recipient(i) & vbCrLf
        Case MAPI_CC:  sCC = sCC & Me.Recipient(i) & vbCrLf
        Case MAPI_BCC: sBC = sCC & Me.Recipient(i) & vbCrLf
        Case Else:   sBad = sBad & Me.Recipient(i) & vbCrLf
        End Select
        If i = 1000 Then Exit For
    Next
    GetRecipients = "To:" & vbCrLf & sTo & _
                    IIf(Len(sCC), vbCrLf & "CC:" & vbCrLf & sCC, "") & _
                    IIf(Len(sBC), vbCrLf & "BCC:" & vbCrLf & sBC, "") & _
                    IIf(Len(sBad), vbCrLf & "bad type:" & vbCrLf & sBad, "")
End Function
