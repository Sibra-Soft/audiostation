VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////
'// FileName        : StringExtensions.cls
'// FileType        : Microsoft Visual Basic 6 - Class Module
'// Type            : Class
'// Author          : Alex van den Berg
'// Created         : 26-06-2023
'// Last Modified   : 14-10-2023
'// Copyright       : Sibra-Soft
'// Description     : Various string functions used in all projects
'////////////////////////////////////////////////////////////////

Option Explicit
Public Function Contains(String1 As String, String2 As String) As Boolean
Contains = (InStr(1, String1, String2) > 0)
End Function
Public Function IsNullOrWhiteSpace(Value As String) As Boolean
If IsNull(Value) Then: IsNullOrWhiteSpace = True
If Value = vbNullString Then: IsNullOrWhiteSpace = True
If Value = " " Then: IsNullOrWhiteSpace = True
If LenB(Value) = 0 Then: IsNullOrWhiteSpace = True
End Function
Public Function Between(First As String, Second As String, Target As String) As String
Dim String1Pos, String2Pos As Integer

String1Pos = InStr(1, Target, First) + 1
String2Pos = InStr(1, Mid(Target, String1Pos), Second) - 1

Between = Mid(Target, String1Pos, String2Pos)
End Function
Public Function format(SourceString As String, ParamArray Values()) As String
Dim i As Integer
Dim ReturnValue As String

ReturnValue = SourceString

For i = 0 To UBound(Values)
    ReturnValue = Replace(ReturnValue, "{" & i & "}", Values(i))
Next

format = ReturnValue
End Function
Public Function StartsWith(ByVal Find As String, ByVal Source As String, Optional ByVal CaseSensitive As Boolean = True) As Boolean
If CaseSensitive Then
    StartsWith = (Left$(Source, LenB(Find) / 2) = Find)
Else
    StartsWith = (Left$(LCase(Source), LenB(Find) / 2) = LCase$(Find))
End If
End Function
Public Function ToCollection(Target As String, Seperator As String) As Collection
Dim i As Integer
Dim SplitValue
Dim TempCollection As New Collection

SplitValue = Split(Target, Seperator)

For i = 0 To UBound(SplitValue)
    TempCollection.Add SplitValue(i)
Next

Set ToCollection = TempCollection
End Function
Public Function SplitStr(TargetStr As String, Delimiter As String, Pos As Integer, Optional Default As String) As String
Dim ReturnValue

ReturnValue = Split(TargetStr, Delimiter)

If UBound(ReturnValue) <= 0 Then
    SplitStr = Default
Else
    SplitStr = ReturnValue(Pos)
End If
End Function
