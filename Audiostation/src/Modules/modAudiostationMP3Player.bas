Attribute VB_Name = "AudiostationMP3Player"
Public PlayState As enumPlayStates
Public PlaylistMode As enumPlaylistMode
Public PlayMode As enumPlayMode
Public PlayStateMediaMode As enumMediaMode

Public Mp3Playlist As New LocalStorage

Public ShowElapsedTime As Boolean
Public CurrentMediaFilename As String
Public CurrentTrackNumber As Integer
Public Sub Init()
PlayState = Stopped
PlaylistMode = RepeatPlaylist
PlayMode = PlaySingleTrack
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
PlayState = Paused
End Sub
Public Sub StartPlay()
AudiostationCDPlayer.StopPlay
AudiostationMIDIPlayer.StopMidiPlayback

PlayStateMediaMode = MP3MediaMode

If PlayState = Paused Then
    Call BASS_ChannelPlay(chan, False)
Else
    If CurrentTrackNumber = 0 Then: CurrentTrackNumber = 1
    
    Dim mediaFilename As String
    
    Call BASS_StreamFree(chan)
    Call BASS_MusicFree(chan)
    
    mediaFilename = Mp3Playlist.GetItemByIndex(CurrentTrackNumber, 1)
    
    CurrentMediaFilename = mediaFilename
    
    chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(mediaFilename), 0, 0, BASS_STREAM_AUTOFREE)
    If chan = 0 Then chan = BASS_MusicLoad(BASSFALSE, mediaFilename, 0, 0, BASS_STREAM_AUTOFREE, 1)
    
    Call BASS_ChannelPlay(chan, True)
End If

PlayState = Playing
End Sub
Public Sub StopPlay()
Call BASS_ChannelStop(chan)
PlayState = Stopped
End Sub
Public Sub NextTrack(Optional TrackNumber As Integer, Optional Force = False)
Dim mediaFilename As String

If Mp3Playlist.StorageContainer.count = 0 Then: Exit Sub
If CurrentTrackNumber = Mp3Playlist.StorageContainer.count Then: Exit Sub

If TrackNumber > 0 Then
    'Track number is set by parameter
    AudiostationMP3Player.CurrentTrackNumber = TrackNumber
Else
    Dim NextTrackNumber As Integer
    Randomize
    
    If Force Then NextTrackNumber = AudiostationMP3Player.CurrentTrackNumber + 1: GoTo DoNext
    
    Select Case AudiostationMP3Player.PlayMode
        Case enumPlayMode.Shuffle: NextTrackNumber = Extensions.RandomNumber(1, Mp3Playlist.StorageContainer.count - 1)
        Case enumPlayMode.PlaySingleTrack: Exit Sub
        Case enumPlayMode.AutoNextTrack
            If PlaylistMode = RepeatSingleTrack Then
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
mediaFilename = Mp3Playlist.GetItemByIndex(CurrentTrackNumber, 1)

CurrentMediaFilename = mediaFilename

Call StartPlay
End Sub
Public Sub PreviousTrack()
Dim mediaFilename As String

If Mp3Playlist.StorageContainer.count = 0 Or CurrentTrackNumber = 1 Then: Exit Sub

AudiostationMP3Player.CurrentTrackNumber = CurrentTrackNumber - 1

mediaFilename = Mp3Playlist.GetItemByIndex(CurrentTrackNumber, 1)
CurrentMediaFilename = mediaFilename

Call StartPlay
End Sub
