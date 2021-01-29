VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmailTextHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Head As String
Private m_Attr_Body As String
Private m_Body As String

Public Sub AddToHead(s As String)
    m_Head = m_Head & s
End Sub
Public Property Get Head() As String
    Dim s As String: s = ""
    s = s & "<head>" & vbCrLf
    s = s & m_Head & vbCrLf
    s = s & "</head>"
    Head = s
End Property

Public Sub AddAttribute_Body(AttrName As String, AttrValue As String)
    m_Attr_Body
End Sub

Public Sub AddToBody(s As String)
    m_Head = m_Head & s
End Sub
Public Property Get body() As String
    Dim s As String: s = ""
    s = s & "<body>" & vbCrLf
    s = s & m_Head & vbCrLf
    s = s & "</body>"
    body = s
End Property

Public Property Get html() As String
    Dim s As String: s = ""
    s = s & "<html>" & vbCrLf
    s = s & Me.Head & vbCrLf
    s = s & Me.body & vbCrLf
    s = s & "</html>" & vbCrLf
    html = s
End Property
