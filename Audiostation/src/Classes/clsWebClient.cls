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

Private Function GenerateQuerystring()
Dim i As Integer
Dim Querystring As String
Dim name As String
Dim Value As String

'For I = 1 To QuerystringParameters.Count
'    name = QuerystringParameters.ItemKey(I)
'    Value = QuerystringParameters.Item(I)
'
'    If I = QuerystringParameters.Count Then
'        Querystring = Querystring & name & "=" & Value
'    Else
'        Querystring = Querystring & name & "=" & Value & "&"
'    End If
'Next

GenerateQuerystring = Querystring
End Function
Public Function WebRequest(url As String) As String
Dim sURL As String
Dim strXML As String
Dim xmlhttp As MSXML2.XMLHTTP60
Dim Querystring As String

Set xmlhttp = New XMLHTTP60

Querystring = GenerateQuerystring()

'If QuerystringParameters.Count > 0 Then URL = URL & "?" & Querystring

xmlhttp.Open "GET", url, False
'xmlhttp.send strXML

WebRequest = xmlhttp.responseText
End Function
