Attribute VB_Name = "AudiostationMP3Player"
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Module:           AudiostationMP3Player
' Description:      Adds MP3 Player functionality
'
' Date Changed:     29-03-2022
' Date Created:     04-10-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public MediaPlaylistMode As enumPlaylistMode
Public MediaPlayMode As enumPlayMode
Public MediaPlaystate As enumPlayStates

Public MediaPlaylist As New LocalStorage

Public ShowElapsedTime As Boolean
Public CurrentMediaFilename As String
Public CurrentTrackNumber As Integer
Public Sub Init()
MediaPlaystate = Stopped
MediaPlaylistMode = RepeatPlaylist
MediaPlayMode = PlaySingleTrack
AudiostationMP3Player.ShowElapsedTime = True
End Sub
Public Sub Rewind()
Dim pos As Long

pos = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan, BASS_POS_BYTE))

Call BASS_ChannelSetPosition(chan, BASS_ChannelSeconds2Bytes(chan, pos + 5), BASS_POS_BYTE)
End Sub
Public Sub Forward()
Dim pos As Long

pos = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan, BASS_POS_BYTE))

Call BASS_ChannelSetPosition(chan, BASS_ChannelSeconds2Bytes(chan, pos - 5), BASS_POS_BYTE)
End Sub
Public Sub Pause()
Call BASS_ChannelPause(chan)
MediaPlaystate = Paused
End Sub
Public Sub StartPlay()
AudiostationCDPlayer.StopPlay
AudiostationMIDIPlayer.StopMidiPlayback

PlayStateMediaMode = MP3MediaMode

If MediaPlaystate = Paused Then
    Call BASS_ChannelPlay(chan, False)
Else
    If CurrentTrackNumber = 0 Then: CurrentTrackNumber = 1
    
    Dim mediaFilename As String
    
    Call BASS_StreamFree(chan)
    Call BASS_MusicFree(chan)
    
    mediaFilename = MediaPlaylist.GetItemByIndex(CurrentTrackNumber, 1)
    
    CurrentMediaFilename = mediaFilename
    
    chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(mediaFilename), 0, 0, BASS_STREAM_AUTOFREE)
    If chan = 0 Then chan = BASS_MusicLoad(BASSFALSE, mediaFilename, 0, 0, BASS_STREAM_AUTOFREE, 1)
    
    Call BASS_ChannelPlay(chan, True)
End If

MediaPlaystate = Playing
End Sub
Public Sub StopPlay()
Call BASS_ChannelStop(chan)
MediaPlaystate = Stopped
End Sub
Public Sub NextTrack(Optional TrackNumber As Integer, Optional Force = False)
Dim mediaFilename As String

If MediaPlaylist.StorageContainer.count = 0 Then: Exit Sub
If CurrentTrackNumber = MediaPlaylist.StorageContainer.count Then: Exit Sub

If TrackNumber > 0 Then
    'Track number is set by parameter
    AudiostationMP3Player.CurrentTrackNumber = TrackNumber
Else
    Dim NextTrackNumber As Integer
    Randomize
    
    If Force Then NextTrackNumber = AudiostationMP3Player.CurrentTrackNumber + 1: GoTo DoNext
    
    Select Case AudiostationMP3Player.MediaPlayMode
        Case enumPlayMode.Shuffle: NextTrackNumber = Extensions.RandomNumber(1, MediaPlaylist.StorageContainer.count - 1)
        Case enumPlayMode.PlaySingleTrack: Exit Sub
        Case enumPlayMode.AutoNextTrack
            If MediaPlaylistMode = RepeatSingleTrack Then
                NextTrackNumber = AudiostationMP3Player.CurrentTrackNumber
            Else
                NextTrackNumber = AudiostationMP3Player.CurrentTrackNumber + 1
            End If
    End Select
    
DoNext:
    'Auto select track number
    AudiostationMP3Player.CurrentTrackNumber = NextTrackNumber
End If

AudiostationMP3Player.CurrentTrackNumber = CurrentTrackNumber
mediaFilename = MediaPlaylist.GetItemByIndex(CurrentTrackNumber, 1)

CurrentMediaFilename = mediaFilename

Call StartPlay
End Sub
Public Sub PreviousTrack()
Dim mediaFilename As String

If MediaPlaylist.StorageContainer.count = 0 Or CurrentTrackNumber = 1 Then: Exit Sub

AudiostationMP3Player.CurrentTrackNumber = CurrentTrackNumber - 1

mediaFilename = MediaPlaylist.GetItemByIndex(CurrentTrackNumber, 1)
CurrentMediaFilename = mediaFilename

Call StartPlay
End Sub
