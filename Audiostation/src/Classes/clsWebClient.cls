VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Class:            WebClient
' Description:      Webrequest class, for getting information from the internet
'
' Date Changed:     05-10-2021
' Date Created:     16-08-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public QuerystringParameters As New Nest
Private Function GenerateQuerystring()
Dim i As Integer
Dim Querystring As String
Dim Name As String
Dim value As String

For i = 1 To QuerystringParameters.count
    Name = QuerystringParameters.ItemKey(i)
    value = QuerystringParameters.Item(i)

    If i = QuerystringParameters.count Then
        Querystring = Querystring & Name & "=" & value
    Else
        Querystring = Querystring & Name & "=" & value & "&"
    End If
Next

GenerateQuerystring = Querystring
End Function
Public Function WebRequest(Url As String) As String
Dim sURL As String
Dim strXML As String
Dim xmlhttp As MSXML2.XMLHTTP60
Dim Querystring As String

Set xmlhttp = New XMLHTTP60

Querystring = GenerateQuerystring()

If QuerystringParameters.count > 0 Then Url = Url & "?" & Querystring

xmlhttp.Open "GET", Url, False
xmlhttp.send strXML

WebRequest = xmlhttp.responseText
End Function
