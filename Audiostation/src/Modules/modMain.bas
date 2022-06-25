Attribute VB_Name = "ModMain"
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Module:           ModMain
' Description:      Main application module of the program
'
' Date Changed:     25-06-2022
' Date Created:     04-10-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public LanguageFile As String

Public Const MAXDRIVES = 10
Public curdrive As Long
Public stream(MAXDRIVES) As Long
Public seeking As Long

Public IsDebuggig As Boolean

Public PlayStateMediaMode As enumMediaMode

Public Mp3Info As New Mp3Info
Public WebRequest As New WebClient
Public Settings As New RegistrySettings
Public Extensions As New SibraSoft

Public AudioStaRecorder As New AudiostationRecorder
Public AudioStaStreamer As New AudiostationSteamer
Public Sub Main()
Call ApplicationConstructor
End Sub
Public Sub OpenFile(MediaFile As String)
Dim MediaIndex As String
Dim MediaDuration As String
Dim CurrentIndex As Integer
Dim CurrentMediaDuration As String

Begin:
If Not Extensions.FileExists(MediaFile) Then Exit Sub

curdrive = Settings.ReadSetting("Sibra-Soft", "Audiostation", "CDDeviceId", 0)

Select Case LCase(right(MediaFile, 3))
    Case "mp3", "wav", "mp2", "aac", "snd", "au", "rmi", "cda", "wma", "m4a"
        MediaDuration = 0
        
        AudiostationMIDIPlayer.StopMidiPlayback
        AudiostationCDPlayer.StopPlay
        
        If MediaPlaylist.IsExistingItem(MediaFile) > 0 Then
            AudiostationMP3Player.CurrentTrackNumber = MediaPlaylist.IsExistingItem(MediaFile)
            AudiostationMP3Player.StartPlay
        Else
            MediaIndex = Format(MediaPlaylist.StorageContainer.count + 1, "00")
            
            ' Only get the duration when it's a mp3 file
            If LCase(right(MediaFile, 3)) = "mp3" Then
                Mp3Info.FileName = MediaFile
                MediaDuration = Extensions.TimeString(Mp3Info.SongLength)
            End If
            
            If MediaDuration = "0" Then: MediaDuration = "-"
            
            MediaPlaylist.AddToStorage MediaFile, MediaIndex & ";" & MediaFile & ";" & MediaDuration
            
            AudiostationMP3Player.CurrentTrackNumber = MediaPlaylist.StorageContainer.count
            AudiostationMP3Player.StartPlay
        End If
                    
    Case "mid", "kar", "mus", "sid"
        AudiostationMP3Player.StopPlay
        AudiostationCDPlayer.StopPlay
        
        CurrentIndex = Format(MidiPlaylist.StorageContainer.count + 1, "00")
        CurrentMediaDuration = "-"

        MidiPlaylist.AddToStorage MediaFile, CurrentIndex & ";" & MediaFile & ";" & CurrentMediaDuration
        
        AudiostationMIDIPlayer.MidiTrackNr = MidiPlaylist.StorageContainer.count
        AudiostationMIDIPlayer.StartMidiPlayback
    
Case "apl", "wpl", "m3u", "pls" 'Playlist files
    If Not (Dir(MediaFile, vbDirectory) = vbNullString) Then
        Screen.MousePointer = vbHourglass
        
        Select Case LCase(right(MediaFile, 3))
            Case "apl": Call ModPlaylist.OpenAplPlaylist(MediaFile)
            Case "m3u": Call ModPlaylist.OpenM3uPlaylist(MediaFile)
            Case "pls": Call ModPlaylist.OpenPlsPlaylist(MediaFile)
            Case "wpl": Call ModPlaylist.OpenWplPlaylist(MediaFile)
        End Select
        
        Form_Playlist.CurrentFormType = Mp3Player
        Form_Playlist.Show , Form_Main
    Else
        Debug.Print "Playlist file could not be found"
    End If
        
Case Else
    'Check if it's a file that needs to be converted
    Select Case LCase(right(MediaFile, 3))
        Case "act": Call ModConvert.Convert(MediaFile, [Voice File Format], MP3): GoTo Begin
        Case "caf": Call ModConvert.Convert(MediaFile, [Apple Core Format], MP3): GoTo Begin
        Case "ogg": Call ModConvert.Convert(MediaFile, [OGG], MP3): GoTo Begin
        Case "omo": Call ModConvert.Convert(MediaFile, [Sony OpenMG Audio], MP3): GoTo Begin
        Case "s64": Call ModConvert.Convert(MediaFile, [Sony Wave64], MP3): GoTo Begin
        Case "voc": Call ModConvert.Convert(MediaFile, [Voice File Format], MP3): GoTo Begin
    End Select
    
    'Check if it's a file that needs to be converted
    Select Case LCase(right(MediaFile, 2))
        Case "ra": Call ModConvert.Convert(MediaFile, [Real Audio], MP3): GoTo Begin
        Case "rm": Call ModConvert.Convert(MediaFile, [Real Media], MP3): GoTo Begin
       
        Case Else: MsgBox GetLanguage(1020), vbCritical
    End Select
End Select
End Sub
Public Sub ApplicationConstructor()
Dim MediaFile As String
Dim MediaIndex As String
Dim MediaDuration As String

If Command = "debugging" Then: IsDebuggig = True

'Check if the temp folder exists
If Dir(App.path & "\temp\", vbDirectory) = vbNullString Then: MkDir (App.path & "\temp\")

'Set the current application langauge
LanguageFile = LCase(Settings.ReadSetting("Sibra-Soft", "Audiostation", "Langauge", "english"))

'Get the loaded file
MediaFile = Command$
MediaFile = Replace(MediaFile, Chr(34), vbNullString)

If Not Extensions.FileExists(MediaFile) Then GoTo Einde
If MediaFile = "" Then: GoTo Einde

If App.PrevInstance Then: Call OpenFile(MediaFile): End

Call OpenFile(MediaFile)
GoTo Einde

Exit Sub

Einde:
    Form_Main.Show
End Sub
Public Sub ApplicationDestructor()
Call BASS_ChannelFree(chan)
Call BASS_Free
End Sub
