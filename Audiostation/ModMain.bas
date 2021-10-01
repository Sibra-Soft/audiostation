Attribute VB_Name = "ModMain"
Public LanguageFile As String

Public Enum enumFormTypes
    [MidiPlayer]
    [Mp3Player]
End Enum

Public IsDebuggig As Boolean
Public WebRequest As New WebClient
Public Settings As New RegistrySettings
Public Extensions As New SibraSoft


Public Sub Main()
Call ApplicationConstructor
End Sub
Public Sub ApplicationConstructor()
Dim MediaFile As String
Dim MediaIndex As String
Dim MediaDuration As String
Dim MediaTagManager As New Mp3Info

If Command = "debugging" Then: IsDebuggig = True

' Clear the Audsta service line
Settings.WriteSetting "Sibra-Soft", "Audiostation", "Audsta", vbNullString

'Check if the temp folder exists
If Dir(App.path & "\temp\", vbDirectory) = vbNullString Then: MkDir (App.path & "\temp\")

'Set the current application langauge
LanguageFile = Settings.ReadSetting("Sibra-Soft", "Audiostation", "Langauge", "english")

'Check if the application is already running
If Settings.ReadSetting("Sibra-Soft", "Audiostation", "ApplicationFirstRun", "1") = 1 And MediaFile = "" Then
    Form_Init.Show
    Exit Sub
Else
    If Settings.ReadSetting("Sibra-Soft", "Audiostation", "ApplicationFirstRun", "1") = 1 And Not MediaFile = "" Then
        Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CheckFile", MediaFile)
        Form_Init.Show
        Exit Sub
    End If
End If

'Get the loaded file
MediaFile = Command$
MediaFile = Replace(MediaFile, Chr(34), vbNullString)

If Not Extensions.FileExists(MediaFile) Then GoTo Einde

If MediaFile = "" Then: GoTo Einde
If App.PrevInstance = True Then: Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CheckFile", MediaFile): End

Select Case LCase(Right(MediaFile, 3))
    Case "mp3", "wav", "mp2", "aac", "snd", "au", "rmi", "cda", "wma", "m4a" 'Audio files
        MediaDuration = 0
        AudiostationMidiPlayer.StopMidiPlayBack
        
        If Mp3Playlist.IsExistingItem(MediaFile) > 0 Then
            AudiostationMP3Player.CurrentTrackNumber = Mp3Playlist.IsExistingItem(MediaFile)
            AudiostationMP3Player.StartPlay
        Else
            CurrentIndex = format(Mp3Playlist.StorageContainer.Count + 1, "00")
            
            If LCase(Right(MediaFile, 3)) = "mp3" Then
                MediaTagManager.FileName = MediaFile
                MediaDuration = Extensions.TimeString(MediaTagManager.SongLength)
            End If
            
            If MediaDuration = "0" Then: MediaDuration = "-"
            
            Mp3Playlist.AddToStorage MediaFile, CurrentIndex & ";" & MediaFile & ";" & MediaDuration
            
            AudiostationMP3Player.CurrentTrackNumber = Mp3Playlist.StorageContainer.Count
            AudiostationMP3Player.StartPlay
        End If
        
    Case "mid", "kar", "mus", "sid" 'Midi files
        AudiostationMP3Player.StopPlay
        
        CurrentIndex = format(MidiPlaylist.StorageContainer.Count + 1, "00")
        CurrentMediaDuration = "-"

        MidiPlaylist.AddToStorage MediaFile, CurrentIndex & ";" & MediaFile & ";" & CurrentMediaDuration
        
        AudiostationMidiPlayer.CurrentMidiTrackNumber = MidiPlaylist.StorageContainer.Count
        AudiostationMidiPlayer.StartMidiPlayback
    
    Case "apl", "wpl", "m3u", "pls" 'Playlist files
        If Not (Dir(MediaFile, vbDirectory) = vbNullString) Then
            Screen.MousePointer = vbHourglass
            
            Select Case LCase(Right(MediaFile, 3))
                Case "apl": Call ModPlaylist.OpenAplPlaylist(MediaFile)
                Case "m3u": Call ModPlaylist.OpenM3uPlaylist(MediaFile)
                Case "pls": Call ModPlaylist.OpenPlsPlaylist(MediaFile)
            End Select
            
            Form_Playlist.CurrentFormType = Mp3Player
            Form_Playlist.Show , Form_Main
        Else
            Debug.Print "Playlist file could not be found"
        End If
        
    Case Else
        'Check if it's a file that needs to be converted
        Select Case LCase(Right(MediaFile, 3))
            Case "act": Call ModConvert.Convert(MediaFile, [Voice File Format], MP3)
            Case "caf": Call ModConvert.Convert(MediaFile, [Apple Core Format], MP3)
            Case "ogg": Call ModConvert.Convert(MediaFile, [OGG], MP3)
            Case "omo": Call ModConvert.Convert(MediaFile, [Sony OpenMG Audio], MP3)
            Case "s64": Call ModConvert.Convert(MediaFile, [Sony Wave64], MP3)
            Case "voc": Call ModConvert.Convert(MediaFile, [Voice File Format], MP3)
        End Select
        
        'Check if it's a file that needs to be converted
        Select Case LCase(Right(file, 2))
            Case "ra": Call ModConvert.Convert(MediaFile, [Real Audio], MP3)
            Case "rm": Call ModConvert.Convert(MediaFile, [Real Media], MP3)
            
            'If the file is not loaded then we can not read it
            Case Else: MsgBox GetLanguage(1020), vbCritical
        End Select
End Select

Einde:
    Form_Main.Show
End Sub
Public Sub ApplicationDestructor()
DoEvents

Call Extensions.TerminateProcessByPid(AudiostationMidiPlayer.CurrentPlayerProcess)

If Not IsDebuggig Then
    Call BASS_WASAPI_Stop(True) ' stop the output
    
    DoEvents
    
    While (BASS_WASAPI_Free)
        DoEvents
    Wend
End If
End Sub
Public Function OutWasapiProc(ByVal Buffer As Long, ByVal Length As Long, ByVal user As Long) As Long
Dim Temp As Long

' check for remaining buffered data
Temp = BASS_WASAPI_GetData(Null, BASS_DATA_AVAILABLE)

If (Temp < 1) Then
    OutWasapiProc = 0
Else
    OutWasapiProc = Temp
End If
End Function
