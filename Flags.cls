VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Flags"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Public Function IconGet(lID As Long) As StdPicture
On Error Resume Next
Set IconGet = LoadResPicture(100 + lID, vbResIcon)
End Function
Public Sub FillCombo(ByVal cbo As Object)
cbo.Clear
Dim i As Integer
For i = 1 To 196
    cbo.AddItem LoadResString(i)
    cbo.ItemData(cbo.NewIndex) = i
Next
End Sub
Public Sub FillImageCombo(ByVal cbo As Object, ByVal iml As Object, Optional lDefIndex As Long = 0)
If TypeName(cbo) <> "ImageCombo" Then Exit Sub
cbo.comboitems.Clear
Dim i As Integer
For i = 1 To 196
    iml.ListImages.Add i, , LoadResPicture(100 + i, vbResIcon)
    cbo.comboitems.Add , "CountryID=" & i, LoadResString(i), i
Next
End Sub
Public Function GetISOCode(ByRef lID As Long) As String
On Error Resume Next
    GetISOCode = LoadResString(200 + lID)
End Function
Public Function GetHTMLCombo(Optional sName As String = "cboCountries", Optional sDefault As String = "") As String
'' Return the HTML combo box code
Dim i As Integer
Dim buf As String
Dim cname As String
On Error Resume Next
buf = "<SELECT name=""" & sName & """ SIZE=""1"">" & vbCrLf
For i = 1 To 196
    cname = LoadResString(i)
    If cname = sDefault Then
        buf = buf & "<OPTION SELECTED VALUE=""" & cname & """>" & cname & "</OPTION>" & vbCrLf
    Else
        buf = buf & "<OPTION VALUE=""" & cname & """>" & cname & "</OPTION>" & vbCrLf
    End If
Next
buf = buf & "</SELECT>" & vbCrLf
GetHTMLCombo = buf
End Function

Public Function GetIDD(ByVal lID As Long) As String
On Error Resume Next
    GetIDD = LoadResString(400 + lID)
End Function
