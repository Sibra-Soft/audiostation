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
Public QuerystringParameters As New Nest
Private Function GenerateQuerystring()
Dim I As Integer
Dim Querystring As String
Dim Name As String
Dim Value As String

For I = 1 To QuerystringParameters.Count
    Name = QuerystringParameters.ItemKey(I)
    Value = QuerystringParameters.Item(I)

    If I = QuerystringParameters.Count Then
        Querystring = Querystring & Name & "=" & Value
    Else
        Querystring = Querystring & Name & "=" & Value & "&"
    End If
Next

GenerateQuerystring = Querystring
End Function
Public Function WebRequest(url As String) As String
Dim sURL As String
Dim strXML As String
Dim xmlhttp As MSXML2.XMLHTTP60
Dim Querystring As String

Set xmlhttp = New XMLHTTP60

Querystring = GenerateQuerystring()

If QuerystringParameters.Count > 0 Then url = url & "?" & Querystring

xmlhttp.Open "GET", url, False
xmlhttp.send strXML

WebRequest = xmlhttp.responseText
End Function
