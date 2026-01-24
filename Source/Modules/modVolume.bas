Attribute VB_Name = "modVolume"
'///////////////////////////////////////////////////////////////
'// FileName        : modVolume.bas
'// FileType        : Microsoft Visual Basic 6 - Module
'// Author          : Alex van den Berg
'// Created         : 04-11-2023
'// Last Modified   : 05-11-2023
'// Copyright       : Sibra-Soft
'// Description     : Audiostation volume channel module
'////////////////////////////////////////////////////////////////

Option Explicit
Public Function ListOfVolChannels() As Collection
Dim returnCollection As New Collection
Dim channelFile As String
Dim fso As New FileSystemObject
Dim SplitValue() As String
Dim I As Integer
Dim ProcessPath As String
Dim ProcessId As Integer
Dim VolumeChannelModel As mdlVolumeChannel

' Remove the old file
If Extensions.FileExists(Environ$("AppData") & "\Audiostation\channels.csv") Then
    Call fso.DeleteFile(Environ$("AppData") & "\Audiostation\channels.csv")
End If

' Generate the new file
Call Shell(App.Path & "\volume.exe /scomma " & Chr(34) & Environ$("AppData") & "\Audiostation\channels.csv" & Chr(34) & " /Columns " & Chr(34) & "Name,Process Path,Process Id" & Chr(34))
Do While Not Extensions.FileExists(Environ$("AppData") & "\Audiostation\channels.csv")
    DoEvents
Loop
channelFile = Extensions.FileGetContents(Environ$("AppData") & "\Audiostation\channels.csv")

' Get the details of the file from the contents
SplitValue = Split(channelFile, vbNewLine)
For I = 1 To UBound(SplitValue)
    ProcessPath = StrExt.SplitStr(SplitValue(I), ",", 1)

    If Not StrExt.IsNullOrWhiteSpace(ProcessPath) Then
        ProcessId = StrExt.SplitStr(SplitValue(I), ",", 2)
        
        Set VolumeChannelModel = New mdlVolumeChannel
        
        VolumeChannelModel.Path = ProcessPath
        VolumeChannelModel.Pid = ProcessId
        VolumeChannelModel.Name = Extensions.GetFileNameFromFilePath(ProcessPath, False)
        
        returnCollection.Add VolumeChannelModel
    End If
Next

Set ListOfVolChannels = returnCollection
End Function
Public Function ChannelExists(AppName As String) As Variant
Dim Channel As mdlVolumeChannel

For Each Channel In ListOfVolChannels
    If StrExt.Contains(LCase$(Channel.Path), LCase$(AppName)) Then
        Set ChannelExists = Channel
        Exit Function
    End If
Next

Set ChannelExists = Nothing
End Function
Public Sub SetVolumeById(Id As String, Value As Integer)
Shell App.Path & "/volume.exe /SetVolume " & Chr(34) & Id & Chr(34) & " " & Value
End Sub
Public Sub SetMuteById(Id As String)
Shell App.Path & "/volume.exe /Mute " & Id
End Sub
Public Sub SetUnmuteById(Id As String)
Shell App.Path & "/volume.exe /Unmute " & Id
End Sub
