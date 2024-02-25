VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////
'// FileName        : Logger.cls
'// FileType        : Microsoft Visual Basic 6 - Class Module
'// Author          : Alex van den Berg
'// Created         : 26-06-2023
'// Last Modified   : 05-11-2023
'// Copyright       : Sibra-Soft
'// Description     : Sibra-Soft application logger class
'////////////////////////////////////////////////////////////////

Option Explicit

Dim FileNum
Public Sub Init(Folder As String)
FileNum = FreeFile

Open Folder & "\audiostation.log" For Append As FileNum
Print #FileNum, Date$ & " " & Time$
End Sub
Public Sub LogError(Message As String)
Print #FileNum, Date$ & vbTab & Time$ & vbTab & "[ERROR]" & vbTab & Message
End Sub
Public Sub LogWarning(Message As String)
Print #FileNum, Date$ & vbTab & Time$ & vbTab & "[WARN]" & vbTab & Message
End Sub
Public Sub LogInfo(Message As String)
Print #FileNum, Date$ & vbTab & Time$ & vbTab & "[INFO]" & vbTab & Message
End Sub
Public Sub CloseLog()
Close FileNum
End Sub

