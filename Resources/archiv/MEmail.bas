Attribute VB_Name = "MEmail"
Option Explicit
''typedef struct {
''  ULONG  ulReserved;
''  ULONG  ulRecipClass;
''  LPSTR  lpszName;
''  LPSTR  lpszAddress;
''  ULONG  ulEIDSize;
''  LPVOID lpEntryID;
''} MapiRecipDesc, *lpMapiRecipDesc;
'Public Type MapiRecip 'Desc
'    ulReserved   As Long   'ULONG
'    ulRecipClass As Long   'ULONG
'    lpszName     As String 'LPSTR
'    lpszAddress  As String 'LPSTR
'    ulEIDSize    As Long   'ULONG
'    lpEntryID    As Long   'LPVOID
'End Type
'
''typedef struct {
''  ULONG  ulReserved;
''  ULONG  flFlags;
''  ULONG  nPosition;
''  LPSTR  lpszPathName;
''  LPSTR  lpszFileName;
''  LPVOID lpFileType;
''} MapiFileDesc, *lpMapiFileDesc;
'Public Type MapiFile 'Desc
'    ulReserved   As Long   'ULONG
'    flFlags      As Long   'ULONG
'    nPosition    As Long   'ULONG
'    lpszPathName As String 'LPSTR
'    lpszFileName As String 'LPSTR
'    lpFileType   As Long   'LPVOID
'End Type
'Const MAPI_OLE        As Long = 0 'The attachment is an OLE object. If MAPI_OLE_STATIC is also set, the attachment is a static OLE object. If MAPI_OLE_STATIC is not set, the attachment is an embedded OLE object.
'Const MAPI_OLE_STATIC As Long = 1 'The attachment is a static OLE object.
'
''typedef struct {
''  ULONG           ulReserved;
''  LPSTR           lpszSubject;
''  LPSTR           lpszNoteText;
''  LPSTR           lpszMessageType;
''  LPSTR           lpszDateReceived;
''  LPSTR           lpszConversationID;
''  FLAGS           flFlags;
''  lpMapiRecipDesc lpOriginator;
''  ULONG           nRecipCount;
''  lpMapiRecipDesc lpRecips;
''  ULONG           nFileCount;
''  lpMapiFileDesc  lpFiles;
''} MapiMessage, *lpMapiMessage;
'Public Type MAPIMessage
'    ulReserved         As Long   'ULONG
'    lpszSubject        As String 'LPSTR
'    lpszNoteText       As String 'LPSTR
'    lpszMessageType    As String 'LPSTR
'    lpszDateReceived   As String 'LPSTR
'    lpszConversationID As String 'LPSTR
'    flFlags            As Long   'FLAGS
'    lpOriginator       As Long   'lpMapiRecipDesc
'    nRecipCount        As Long   'ULONG
'    lpRecips           As Long   'lpMapiRecipDesc
'    nFileCount         As Long   'ULONG
'    lpFiles            As Long   'lpMapiFileDesc
'End Type
'
'
'Const MAPI_RECEIPT_REQUESTED As Long = 0 'A receipt notification is requested. Client applications set this flag when sending a message.
'Const MAPI_SENT              As Long = 1 'The message has been sent.
'Const MAPI_UNREAD            As Long = 2 'The message has not been read.
'
'Public Declare Function MAPISendMail Lib "MAPI32.DLL" Alias "BMAPISendMail" ( _
'    ByVal Session As Long, _
'    ByVal UIParam As Long, _
'    ByRef Message As MAPIMessage, _
'    ByRef Recipient() As MapiRecip, _
'    ByRef File() As MapiFile, _
'    ByVal flags As Long, _
'    ByVal Reserved As Long) As Long

'-----------------------------------------------------------------
' Procedure : SendToMailRecipient
' Purpose   : Simulates a drop operation to
'             "Sent To/Mail Recipient" shell extension
'-----------------------------------------------------------------
'
'Public Sub SendToMailRecipient(ByVal FileName As String)
'
'    ' Initialize interface of IDropTarget
'    Dim tIID_IDropTarget As UUID: CLSIDFromString "{00000122-0000-0000-C000-000000000046}", tIID_IDropTarget
'
'    ' Initialize CLSID of ".MAPIMail"
'    Dim tCLSID_SendMail  As UUID: CLSIDFromString "{9E56BE60-C50F-11CF-9A2C-00A0C90A90CE}", tCLSID_SendMail
'
'    Dim lRes As Long
'    ' Create the "SendTo/Mail Recipient" object
'    Dim oSendMail        As IDropTarget: lRes = CoCreateInstance(tCLSID_SendMail, Nothing, CLSCTX_INPROC_SERVER, tIID_IDropTarget, oSendMail)
'
'    If lRes = S_OK Then
'        ' Get the file IDataObject interface
'        Dim oDO As IDataObject: Set oDO = GetFileDataObject(FileName)
'        ' Simulate the drop operation
'        oSendMail.DragEnter oDO, vbKeyLButton, 0, 0, DROPEFFECT_COPY
'        oSendMail.Drop oDO, vbKeyLButton, 0, 0, DROPEFFECT_COPY
'        oSendMail
'    Else
'        Err.Raise lRes
'    End If
'End Sub

'--------------------------------------------------------------
' Procedure : GetFileDataObject
' Purpose   : Returns the IDataObject interface for a file
'--------------------------------------------------------------
'
'Private Function GetFileDataObject(ByVal FileName As String) As IDataObject
'    ' Intialize IDs
'    Dim tIID_IDataObject  As UUID: CLSIDFromString "{0000010e-0000-0000-C000-000000000046}", tIID_IDataObject
'    Dim tIID_IShellFolder As UUID: CLSIDFromString IIDSTR_IShellFolder, tIID_IShellFolder
'    Dim sFolder As String: sFolder = Left$(FileName, InStrRev(FileName, "\") - 1)
'    FileName = Mid$(FileName, Len(sFolder) + 2)
'    If Right$(sFolder, 1) = ":" Then sFolder = sFolder + "\"
'    ' Get the parent folder object
'    Dim oDesktop As IShellFolder: Set oDesktop = SHGetDesktopFolder
'    ' Get the parent folder IDL
'    Dim lPidl As Long, lPtr As Long
'    oDesktop.ParseDisplayName 0, 0, StrPtr(sFolder), lPtr, lPidl, 0
'    ' Get the parent folder object
'    oDesktop.BindToObject lPidl, 0, tIID_IShellFolder, lPtr
'    Dim oParent As IShellFolder: MoveMemory oParent, lPtr, 4&
'    ' Release the PIDL
'    CoTaskMemFree lPidl
'    ' Get the file PIDL
'    oParent.ParseDisplayName 0, 0, StrPtr(FileName), 0, lPidl, 0
'    ' Get the file IDataObject
'    lPtr = oParent.GetUIObjectOf(0, 1, lPidl, tIID_IDataObject, 0)
'    Dim oUnk As IUnknown: MoveMemory oUnk, lPtr, 4&
'    ' Release the file PIDL
'    CoTaskMemFree lPidl
'    ' Return the file IDataObject
'    Set GetFileDataObject = oUnk
'End Function
