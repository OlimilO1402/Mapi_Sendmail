VERSION 5.00
Begin VB.Form FrmEmail 
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8415
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnEmailStart 
      Caption         =   "Start Email"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      ToolTipText     =   "No email will be sent, unless you click <send> in your emailprogram!"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton BtnFillTestEmail 
      Caption         =   "Fill Test-Email"
      Height          =   375
      Left            =   840
      TabIndex        =   12
      ToolTipText     =   "Fille some example data in this dialog."
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton BtnEmailLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.ListBox LBFiles 
      Height          =   2010
      ItemData        =   "Form1.frx":1782
      Left            =   4800
      List            =   "Form1.frx":1784
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   1
      ToolTipText     =   "drag'drop files here"
      Top             =   0
      Width           =   3615
   End
   Begin VB.TextBox TxtBodyText 
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   10
      Top             =   2640
      Width           =   8415
   End
   Begin VB.TextBox TxtSubject 
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   2160
      Width           =   7575
   End
   Begin VB.TextBox TxtRecpBCC 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox TxtRecpCC 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox TxtRecpTo 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "Files:"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "BCC:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "CC:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "FrmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Email As StdMail

Private Sub Form_Load()
    Set m_Email = New StdMail
End Sub
Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    L = LBFiles.Left: W = Me.ScaleWidth - L: H = LBFiles.Height
    If W > 0 And H > 0 Then LBFiles.Move L, T, W, H
    L = 0: T = TxtBodyText.Top: W = Me.ScaleWidth: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then TxtBodyText.Move L, T, W, H
End Sub

Private Sub BtnFillTestEmail_Click()
    TxtRecpTo.Text = "Hugh man musterman <human@musterman.net>"
    TxtRecpCC.Text = "Anotherone Bitesthedust <anotherone@bitesthedust.com>"
    TxtSubject.Text = "Hello for Test"
    TxtBodyText.Text = GetHtmlEmail.ToHtmlStr
    LBFiles.AddItem App.Path & "\Resources\Test1.pdf"
    LBFiles.AddItem App.Path & "\Resources\Test2.pdf"
End Sub

'Private Sub Command4_Click()
'    Debug.Print GetHtmlEmail.ToHtmlStr
'End Sub

Private Sub BtnEmailStart_Click()
    Set m_Email = New StdMail
    With m_Email
        .IsUnicode = True
        .RecipientAddTo TxtRecpTo.Text
        .RecipientAddCC TxtRecpCC.Text
        .Subject = TxtSubject.Text '"Hello for Test"
        .BodyText = TxtBodyText.Text 'GetHtmlEmail.ToHtmlStr
        Dim i As Long
        For i = 0 To LBFiles.ListCount - 1
            .FileAdd LBFiles.List(i)
        Next
        '.FileAdd App.Path & "\Resources\Test1.pdf"
        '.FileAdd App.Path & "\Resources\Test2.pdf"
    End With
    m_Email.Start
End Sub
Private Function AddRecipientsTo(TB As TextBox, eml As StdMail)
    '
End Function
Private Sub BtnEmailLogin_Click()
    m_Email.Login "XXXXXXXXXXXXXXXXXXXX", "XXXXXXXXXXXXXXXXXXXX"
End Sub
Private Sub LBFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFFiles) Then
        Dim f
        For Each f In Data.Files
            LBFiles.AddItem f 'GetFileName(f)
        Next
    End If
End Sub
Function GetFileName(ByVal aPFN As String) As String
    Dim pos As Long: pos = InStrRev(aPFN, "\")
    If pos = 0 Then GetFileName = aPFN: Exit Function
    GetFileName = Right(aPFN, Len(aPFN) - pos)
End Function

Private Sub TxtRecpTo_LostFocus()
    'Ja, wie das parsen und gleichzeitig schauen ob schon enthalten ist,
    'soll man direkt in die Email schreiben, ja wegen SessionID muss man direkt in die Email schreiben
    'd.h
End Sub
Public Function getHTMLText() As String
    Dim s As String: s = ""
    s = s & "<html>" & vbCrLf
    s = s & "  <head>" & vbCrLf
    s = s & "  </head>" & vbCrLf
    s = s & "  <body bgcolor=""#FFFFFF"" style=""font-family: Verdana;font-size: 12.0px;"" text=""#000000"">" & vbCrLf
    s = s & "    <p>Hallo,</p>" & vbCrLf
    s = s & "    <blockquote>" & vbCrLf
    s = s & "      <p>" & vbCrLf
    s = s & "        <u>" & vbCrLf
    s = s & "          <i>" & vbCrLf
    s = s & "            <b>" & vbCrLf
    s = s & "              <font size=""+2"">dies ist eine Email im html-Format.<span class=""moz-smiley-s1""><span>:-)</span></span><br/>" & vbCrLf
    s = s & "              </font>"
    s = s & "            </b>"
    s = s & "          </i>"
    s = s & "        </u>"
    s = s & "      </p>" & vbCrLf
    s = s & "      <hr size=""2"" width=""100%""/>"
    s = s & "    </blockquote>" & vbCrLf
    s = s & "    <hr size=""2"" width=""100%""/>" & vbCrLf
    s = s & "    <table border=""1"" cellpadding=""2"" cellspacing=""2"" width=""100%"">" & vbCrLf
    s = s & "      <tbody>" & vbCrLf
    s = s & "        <tr>" & vbCrLf
    s = s & "          <td valign=""top"">1<br/>" & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "          <td valign=""top"">2<br/>" & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "          <td valign=""top"">3<br/>" & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "        </tr>" & vbCrLf
    s = s & "        <tr>" & vbCrLf
    s = s & "          <td valign=""top"">4<br/>" & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "          <td valign=""top"">5<br/>" & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "          <td valign=""top"">6<br/>" & vbCrLf
    s = s & "          </td>" & vbCrLf
    s = s & "        </tr>" & vbCrLf
    s = s & "      </tbody>" & vbCrLf
    s = s & "    </table>" & vbCrLf
    s = s & "    <p><br/>" & vbCrLf
    s = s & "    </p>" & vbCrLf
    s = s & "    <p>Gru&szlig;</p>" & vbCrLf
    s = s & "    <p>Oliver<br/>" & vbCrLf
    s = s & "    </p>" & vbCrLf
    s = s & "  </body>" & vbCrLf
    s = s & "</html>" & vbCrLf
    
    getHTMLText = s
End Function


Public Function GetHtmlEmail() As HtmlElem
    Dim html As HtmlElem
    Set html = New_HtmlElem("html")
    Dim body As HtmlElem
    With html
        .AddElem "head"
        With .AddElem("body").AddAttr("bgcolor", "#FFFFFF").AddAttr("style", "font-family: Verdana; font-size: 12.0px;").AddAttr("text", "#000000")
            With .AddElem("p").SetInnText("Hallo,")
                With .AddElem("blockquote")
                    With .AddElem("p"): With .AddElem("u"): With .AddElem("i"): With .AddElem("b")
                                    With .AddElem("font").AddAttr("size", "+2")
                                        .SetInnText "dies ist eine Email im html-Format."
                                        With .AddElem("span").AddAttr("class", "moz-smiley-s1")
                                            .AddElem("span").SetInnText ":-)"
                                        End With
                                        .AddElem("br").SetEnd True
                                    End With
                                End With
                            End With
                        End With
                    End With
                    .AddElem("hr").AddAttr("size", "2").AddAttr("width", "100%").SetEnd True
                End With
                .AddElem("hr").AddAttr("size", "2").AddAttr("width", "100%").SetEnd True
                With .AddElem("table").AddAttr("border", "1").AddAttr("cellpadding", "2").AddAttr("cellspacing", "2").AddAttr("width", "100%")
                    With .AddElem("tbody")
                        With .AddElem("tr")
                            With .AddElem("td").AddAttr("valign", "top").SetInnText("1"): .AddElem("br").SetEnd True: End With
                            With .AddElem("td").AddAttr("valign", "top").SetInnText("2"): .AddElem("br").SetEnd True: End With
                            With .AddElem("td").AddAttr("valign", "top").SetInnText("3"): .AddElem("br").SetEnd True: End With
                        End With
                        With .AddElem("tr")
                            With .AddElem("td").AddAttr("valign", "top").SetInnText("4"): .AddElem("br").SetEnd True: End With
                            With .AddElem("td").AddAttr("valign", "top").SetInnText("5"): .AddElem("br").SetEnd True: End With
                            With .AddElem("td").AddAttr("valign", "top").SetInnText("6"): .AddElem("br").SetEnd True: End With
                        End With
                    End With
                End With
                With .AddElem("p"): .AddElem("br").SetEnd True: End With
                .AddElem("p").SetInnText "Gru&szlig;"
                With .AddElem("p")
                    .SetInnText "Oliver"
                    .AddElem("br").SetEnd True
                End With
            End With
        End With
    End With
    Set GetHtmlEmail = html
End Function
'Sub MapiSimpleSendMail(Files() As String, ToEmailAddress As String, toName As String)
'
'    Dim aMAPIMessage As MAPIMessage
'    Dim Flags As Long
'    Dim emailSubject As String
'    Dim emailBody    As String
'    Dim recipients()   As MapiRecip
'    Dim attachements() As MapiFile
'    Dim i  As Long
'    Dim hr As Long
'    Dim es As String
'    Const MAPI_E_UNICODE_NOT_SUPPORTED As Long = 27
'    aMAPIMessage.lpOriginator = 0
'    aMAPIMessage.Subject = StrConv("Achtung Datei", vbFromUnicode)
'    If ToEmailAddress <> "" Then
'        recipients(0).RecipClass = MAPI_TO
'        recipients(0).Address = StrConv(ToEmailAddress, vbFromUnicode)
'        recipients(0).Name = StrConv(toName, vbFromUnicode)
'        aMAPIMessage.lpRecips = VarPtr(recipients(0))
'        aMAPIMessage.nRecipCount = 1
'    Else
'        aMAPIMessage.lpRecips = 0
'        aMAPIMessage.nRecipCount = 0
'    End If
'    aMAPIMessage.MessageType = 0
'    If UBound(Files) > 0 Then
'        emailSubject = "Emailing:"
'        '//Yes, the shell really does create a blank mail with a leading line of ten spaces
'        emailBody = "          " & vbCrLf & _
'                    "Your message is ready to be sent with the following file or link attachments:" & vbCrLf
'        ReDim attachements(0 To UBound(Files))
'        For i = 0 To UBound(Files)
'            attachements(i).FileName = StrConv(Files(i), vbFromUnicode)
'            attachements(i).nPosition = &HFFFFFFFF
'        Next
'    Else
'        '
'    End If
'    aMAPIMessage.Subject = StrConv(emailSubject, vbFromUnicode)
'    aMAPIMessage.NoteText = StrConv(emailBody, vbFromUnicode)
'    Flags = MAPI_DIALOG
'
'    hr = MAPISendMail(0, 0, aMAPIMessage, Flags, 0)
'    Debug.Print hr
'
'End Sub
'
'procedure MapiSimpleSendMail(slFiles: TStrings; ToEmailAddress: string=''; ToName: string='');
'Var
'    mapiMessage: TMapiMessage;
'    flags: LongWord;
'//  senderName: AnsiString;
'//  senderEmailAddress: AnsiString;
'    emailSubject: AnsiString;
'    emailBody: AnsiString;
'//  sender: TMapiRecipDesc;
'    recipients: packed array of TMapiRecipDesc;
'    attachments: packed array of TMapiFileDesc;
'    i: Integer;
'    hr: Cardinal;
'    es: string;
'const
'    MAPI_E_UNICODE_NOT_SUPPORTED = 27; //Windows 8. The MAPI_FORCE_UNICODE flag is specified and Unicode is not supported.
'Begin
'    ZeroMemory(@mapiMessage, SizeOf(mapiMessage));
'
'{   senderName := '';
'    senderEmailAddress := '';
'
'    ZeroMemory(@sender, sizeof(sender));
'    sender.ulRecipClass := MAPI_ORIG; //MAPI_TO, MAPI_CC, MAPI_BCC, MAPI_ORIG
'    sender.lpszName := PAnsiChar(senderName);
'    sender.lpszAddress := PAnsiChar(senderEmailAddress);
'}
'
'    mapiMessage.lpOriginator := nil; //PMapiRecipDesc; { Originator descriptor                  }
'
'    if ToEmailAddress <> '' then
'    Begin
'        SetLength(recipients, 1);
'        recipients[0].ulRecipClass := MAPI_TO;
'        recipients[0].lpszName := LPSTR(ToName);
'        recipients[0].lpszAddress := LPSTR(ToEmailAddress);
'
'        mapiMessage.lpRecips := @recipients[0]; //A value of NULL means that there are no recipients. Additionally, when this member is NULL, the nRecipCount member must be zero.
'        mapiMessage.nRecipCount := 1;
'    End
'    Else
'    Begin
'        mapiMessage.lpRecips := nil; //A value of NULL means that there are no recipients. Additionally, when this member is NULL, the nRecipCount member must be zero.
'        mapiMessage.nRecipCount := 0;
'    end;
'
'    mapiMessage.lpszMessageType := nil;
'
'    If slFiles.Count > 0 Then
'    Begin
'        emailSubject := 'Emailing: ';
'        emailBody :=
'                '          '+#13#10+ //Yes, the shell really does create a blank mail with a leading line of ten spaces
'                'Your message is ready to be sent with the following file or link attachments:'+#13#10;
'
'
'        SetLength(attachments, slFiles.Count);
'        for i := 0 to slFiles.Count-1 do
'        Begin
'            attachments[i].ulReserved := 0; // Cardinal;        { Reserved for future use (must be 0)     }
'            attachments[i].flFlags := 0; // Cardinal;           { Flags                                   }
'            attachments[i].nPosition := $FFFFFFFF; //Cardinal;         { character in text to be replaced by attachment }
'            attachments[i].lpszPathName := PAnsiChar(slFiles[i]);    { Full path name of attachment file       }
'            attachments[i].lpszFileName := nil; // LPSTR;         { Original file name (optional)           }
'            attachments[i].lpFileType := nil; // Pointer;         { Attachment file type (can be lpMapiFileTagExt) }
'
'            If i > 0 Then
'                emailSubject := emailSubject+', ';
'            emailSubject := emailSubject+ExtractFileName(slFiles[i]);
'            emailBody := emailBody+#13#10+
'                    ExtractFileName(slFiles[i]);
'        end;
'
'        emailBody := emailBody+#13#10+
'                #13#10+
'                #13#10+
'                'Note: To protect against computer viruses, e-mail programs may prevent sending or receiving certain types of file attachments.  Check your e-mail security settings to determine how attachments are handled.';
'
'
'        mapiMessage.lpFiles := @attachments[0];
'        mapiMessage.nFileCount := slFiles.Count;
'    End
'    Else
'    Begin
'        emailSubject := '';
'        emailBody := '';
'
'        mapiMessage.lpFiles := nil;
'        mapiMessage.nFileCount := 0;
'    end;
'
'    {
'        Subject
'        Emailing: 4388_888871544_MVM_10.tmp, amt3.log, swtag.log, wct845C.tmp, ~vs1830.sql
'
'        Body
'                  <-- ten spaces
'        Your message is ready to be sent with the following file or link attachments:
'
'        4388_888871544_MVM_10.tmp
'        amt3.Log
'        swtag.Log
'        wct845C.tmp
'        ~vs1830.sql
'
'
'        Note: To protect against computer viruses, e-mail programs may prevent sending or receiving certain types of file attachments.  Check your e-mail security settings to determine how attachments are handled.
'    }
'    mapiMessage.lpszSubject := PAnsiChar(emailSubject);
'    mapiMessage.lpszNoteText := PAnsiChar(emailBody);
'
'
'    flags := MAPI_DIALOG;
'
'    hr := Mapi.MapiSendMail(0, 0, mapiMessage, flags, 0);
'    case hr of
'    SUCCESS_SUCCESS: {nop}; //The call succeeded and the message was sent.
'MAPI_E_AMBIGUOUS_RECIPIENT:
'        Begin
'            //es := 'A recipient matched more than one of the recipient descriptor structures and MAPI_DIALOG was not set. No message was sent.';
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_AMBIGUOUS_RECIPIENT', SysErrorMessage(hr)]);
'        end;
'MAPI_E_ATTACHMENT_NOT_FOUND:
'        Begin
'            //The specified attachment was not found. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_ATTACHMENT_NOT_FOUND', SysErrorMessage(hr)]);
'        end;
'MAPI_E_ATTACHMENT_OPEN_FAILURE:
'        Begin
'            //The specified attachment could not be opened. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_ATTACHMENT_OPEN_FAILURE', SysErrorMessage(hr)]);
'        end;
'MAPI_E_BAD_RECIPTYPE:
'        Begin
'            //The type of a recipient was not MAPI_TO, MAPI_CC, or MAPI_BCC. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_BAD_RECIPTYPE', SysErrorMessage(hr)]);
'        end;
'MAPI_E_FAILURE:
'        Begin
'            //One or more unspecified errors occurred. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_FAILURE', SysErrorMessage(hr)]);
'        end;
'MAPI_E_INSUFFICIENT_MEMORY:
'        Begin
'            //There was insufficient memory to proceed. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_INSUFFICIENT_MEMORY', SysErrorMessage(hr)]);
'        end;
'MAPI_E_INVALID_RECIPS:
'        Begin
'            //One or more recipients were invalid or did not resolve to any address.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_INVALID_RECIPS', SysErrorMessage(hr)]);
'        end;
'MAPI_E_LOGIN_FAILURE:
'        Begin
'            //There was no default logon, and the user failed to log on successfully when the logon dialog box was displayed. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_LOGIN_FAILURE', SysErrorMessage(hr)]);
'        end;
'MAPI_E_TEXT_TOO_LARGE:
'        Begin
'            //The text in the message was too large. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_TEXT_TOO_LARGE', SysErrorMessage(hr)]);
'        end;
'MAPI_E_TOO_MANY_FILES:
'        Begin
'            //There were too many file attachments. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_TOO_MANY_FILES', SysErrorMessage(hr)]);
'        end;
'MAPI_E_TOO_MANY_RECIPIENTS:
'        Begin
'            //There were too many recipients. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_TOO_MANY_RECIPIENTS', SysErrorMessage(hr)]);
'        end;
'MAPI_E_UNICODE_NOT_SUPPORTED:
'        Begin
'            //The MAPI_FORCE_UNICODE flag is specified and Unicode is not supported.
'            //Note  This value can be returned by MAPISendMailW only.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_UNICODE_NOT_SUPPORTED', SysErrorMessage(hr)]);
'        end;
'MAPI_E_UNKNOWN_RECIPIENT:
'        Begin
'            //A recipient did not appear in the address list. No message was sent.
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_UNKNOWN_RECIPIENT', SysErrorMessage(hr)]);
'        end;
'MAPI_E_USER_ABORT:
'        Begin
'            es := 'The user canceled one of the dialog boxes. No message was sent.';
'            raise Exception.CreateFmt('Error %s sending e-mail message: %s', ['MAPI_E_USER_ABORT', es]);
'        end;
'    Else
'        raise Exception.CreateFmt('Error %d sending e-mail message: %s', [hr, SysErrorMessage(hr)]);
'    end;
'end;

