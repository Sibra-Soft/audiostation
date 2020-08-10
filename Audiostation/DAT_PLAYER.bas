Attribute VB_Name = "AudiostationMP3Player"
Public Mp3Playlist As New LocalStorage

Public RepeatTrack As Boolean
Public RepeatPlaylist As Boolean
Public AutoNext As Boolean
Public Shuffle As Boolean
Public PlaySingleTrack As Boolean
Public ShowElapsedTime As Boolean
Public Stopped As Boolean

Public CurrentMediaFilename As String

Public CurrentTrackNumber As Integer
Public Sub Init()
AudiostationMP3Player.ShowElapsedTime = True
End Sub
Public Sub Rewind()
Dim newPosition As TStreamTime

newPosition.sec = 5

Call ModLibZPlay.SeekPosition(tfSecond, newPosition, smFromCurrentForward)
End Sub
Public Sub Forward()
Dim newPosition As TStreamTime

newPosition.sec = 5

Call ModLibZPlay.SeekPosition(tfSecond, newPosition, smFromCurrentBackward)
End Sub
Public Sub Pause()
ModLibZPlay.PausePlayback
Stopped = True
End Sub
Public Sub StartPlay()
Dim mediaFilename As String
Dim StreamStatus As TStreamStatus

Call ModLibZPlay.GetStatus(StreamStatus)

AudiostationMidiPlayer.StopMidiPlayBack

If StreamStatus.fPause Then
    Call ModLibZPlay.ResumePlayback
Else
    If CurrentTrackNumber = 0 Then: CurrentTrackNumber = 1
    
    mediaFilename = Mp3Playlist.GetItemByIndex(CurrentTrackNumber, 1)
    CurrentMediaFilename = mediaFilename
        
    Call ModLibZPlay.OpenFile(mediaFilename, sfAutodetect)
    Call ModLibZPlay.StartPlayback
End If

Stopped = False
End Sub
Public Sub StopPlay()
ModLibZPlay.StopPlayback
Stopped = True
End Sub
Public Sub nextTrack(Optional TrackNumber As Integer)
Dim mediaFilename As String

If Mp3Playlist.StorageContainer.Count = 0 Then: Exit Sub
If AudiostationMP3Player.CurrentTrackNumber = Mp3Playlist.StorageContainer.Count Then: Exit Sub

If TrackNumber > 0 Then
    'Track number is set by parameter
    AudiostationMP3Player.CurrentTrackNumber = TrackNumber
Else
    'Auto select track number
    AudiostationMP3Player.CurrentTrackNumber = AudiostationMP3Player.CurrentTrackNumber + 1
End If

AudiostationMP3Player.CurrentTrackNumber = CurrentTrackNumber
mediaFilename = Mp3Playlist.GetItemByIndex(CurrentTrackNumber, 1)

CurrentMediaFilename = mediaFilename

Call ModLibZPlay.OpenFile(mediaFilename, sfAutodetect)
Call ModLibZPlay.StartPlayback

Stopped = False
End Sub
Public Sub PreviousTrack()
Dim mediaFilename As String

If Mp3Playlist.StorageContainer.Count = 0 Or CurrentTrackNumber = 1 Then: Exit Sub

AudiostationMP3Player.CurrentTrackNumber = CurrentTrackNumber - 1

mediaFilename = Mp3Playlist.GetItemByIndex(CurrentTrackNumber, 1)
CurrentMediaFilename = mediaFilename

Call ModLibZPlay.OpenFile(mediaFilename, sfAutodetect)
Call ModLibZPlay.StartPlayback

Stopped = False
End Sub
