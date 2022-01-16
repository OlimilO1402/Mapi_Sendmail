Attribute VB_Name = "MNew"
Option Explicit

Sub Main()
    FMain.Show
End Sub

Public Function New_HtmlElem(aName As String) As HtmlElem
    Set New_HtmlElem = New HtmlElem
    New_HtmlElem.New_ aName
End Function

