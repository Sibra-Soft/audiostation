Attribute VB_Name = "AudiostationRecorder"
Public RecordActive As Boolean
Public RecordFilename As String
Public Sub StartRecorder()
Dim SPResult As SP_RESULTS

SPResult = Form_Main.ShellPipe.Run(Chr(34) & App.path & "\ffmpeg.exe" & Chr(34) & " -f dshow -i audio=" & Chr(34) & "virtual-audio-capturer" & Chr(34) & " " & Chr(34) & RecordFilename & Chr(34), App.path)

Select Case SPResult
    Case SP_CREATEPIPEFAILED: MsgBox "Run failed, could not create pipe", vbOKOnly Or vbExclamation, Caption
    Case SP_CREATEPROCFAILED: MsgBox "Run failed, could not create process", vbOKOnly Or vbExclamation, Caption
End Select

AudiostationRecorder.RecordActive = True
End Sub
Public Sub StopRecorder()
Form_Main.ShellPipe.SendData "q"

AudiostationRecorder.RecordActive = False

Call SaveRecordFile
End Sub
Public Sub SaveRecordFile()
Dim fso As New FileSystemObject

On Error GoTo ErrorHandler
If Not Extensions.FileExists(RecordFilename) Then: MsgBox GetLanguage(1025), vbCritical: Exit Sub

If Extensions.INIRead("wav", "WavOutputLocation", App.path & "\settings.ini") = vbNullString Then
    With Form_Main
        .CommonDialog1.CancelError = True
        .CommonDialog1.DialogTitle = GetLanguage(1026)
        .CommonDialog1.Filter = "Microsoft WaveForm Audio (*.wav)|*.wav"
        .CommonDialog1.ShowSave
                    
        fso.CopyFile RecordFilename, .CommonDialog1.FileName, True
        
        Extensions.Pause 200
        
        Kill RecordFilename
    End With
End If

ErrorHandler:
Select Case err.Number
    Case 0 ' Do nothing
    Case cdlCancel ' Do nothing
End Select
End Sub
