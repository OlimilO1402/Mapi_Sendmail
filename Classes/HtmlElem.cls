VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HtmlElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Name    As String
Private m_AttrNam As Collection
Private m_AttrVal As Collection
Private m_InnElem As Collection 'Of HtmlElement
Private m_InnText As String
Private m_bEnd    As Boolean

'Private Sub Class_Initialize()
    'Set m_AttrNam = New Collection
    'Set m_AttrVal = New Collection
    'Set m_InnElem = New Collection
'End Sub

Public Sub New_(aName As String)
    m_Name = aName
End Sub
Public Function SetInnText(s As String) As HtmlElem
    m_InnText = s
    Set SetInnText = Me
End Function
Public Function AddAttr(aName As String, aValue As String) As HtmlElem
    If m_AttrNam Is Nothing Then
        Set m_AttrNam = New Collection
        Set m_AttrVal = New Collection
    End If
    m_AttrNam.Add aName
    m_AttrVal.Add aValue, aName
    Set AddAttr = Me
End Function

'Public Function AddElem(e As HtmlElem) As HtmlElem
Public Function AddElem(eName As String) As HtmlElem
    If m_InnElem Is Nothing Then Set m_InnElem = New Collection
    Dim elem As New HtmlElem: elem.New_ eName
    m_InnElem.Add elem
    Set AddElem = elem
    'Set AddElem = Me
End Function
Public Function SetEnd(ByVal bEnd As Boolean) As HtmlElem
    m_bEnd = bEnd
    Set SetEnd = Me
End Function
Private Function getAttr() As String
    If m_AttrNam Is Nothing Then Exit Function
    Dim s As String
    Dim v, sn As String
    For Each v In m_AttrNam
        sn = v
        s = s & " " & sn & "=" & """" & m_AttrVal(sn) & """"
    Next
    getAttr = s
End Function
Private Function getElem() As String
    If m_InnElem Is Nothing Then Exit Function
    Dim s As String
    Dim he As HtmlElem
    For Each he In m_InnElem
        s = s & he.ToHtmlStr
    Next
    getElem = s
End Function
Public Function ToHtmlStr() As String
    Dim s As String
    s = s & "<" & m_Name & getAttr
    If m_bEnd Then
        s = s & "/" & ">" '& vbCrLf
    Else
        s = s & ">" '& vbCrLf
        s = s & m_InnText
        s = s & getElem '& vbCrLf
        s = s & "</" & m_Name & ">" & vbCrLf
    End If
    ToHtmlStr = s
End Function

