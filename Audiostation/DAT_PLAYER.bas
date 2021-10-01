Attribute VB_Name = "AudiostationMP3Player"
Public Enum EnumPlayStates
    Paused
    Stopped
    Playing
    MediaEnded
End Enum

Public MediaPlayer As New MediaPlayer
Public MediaPlayerForCD As New MediaPlayerForCD

Public Mp3Playlist As New LocalStorage

Public RepeatTrack As Boolean
Public RepeatPlaylist As Boolean
Public AutoNext As Boolean
Public Shuffle As Boolean
Public PlaySingleTrack As Boolean
Public ShowElapsedTime As Boolean

Public PlayState As EnumPlayStates
Public CurrentMediaFilename As String

Public CurrentTrackNumber As Integer
Public Sub Init()
PlayState = Stopped
AudiostationMP3Player.ShowElapsedTime = True
End Sub
Public Sub Rewind()
Dim CurrentPosition As Long

CurrentPosition = MediaPlayer.GetPositioninSec
MediaPlayer.ChangePosition CurrentPosition + 5
End Sub
Public Sub Forward()
Dim CurrentPosition As Long

CurrentPosition = MediaPlayer.GetPositioninSec
MediaPlayer.ChangePosition CurrentPosition - 5
End Sub
Public Sub Pause()
MediaPlayer.Pause
PlayState = Paused
End Sub
Public Sub StartPlay()
AudiostationMidiPlayer.StopMidiPlayBack

If PlayState = Paused Then
    MediaPlayer.ResumePlay
Else
    If CurrentTrackNumber = 0 Then: CurrentTrackNumber = 1
    
    Dim mediaFilename As String
    mediaFilename = Mp3Playlist.GetItemByIndex(CurrentTrackNumber, 1)
    
    CurrentMediaFilename = mediaFilename
    
    MediaPlayer.FileName = mediaFilename
    MediaPlayer.Play
End If

PlayState = Playing
End Sub
Public Sub StopPlay()
MediaPlayer.StopPlay
PlayState = Stopped
End Sub
Public Sub NextTrack(Optional TrackNumber As Integer, Optional force = False)
Dim mediaFilename As String

If Mp3Playlist.StorageContainer.Count = 0 Then: Exit Sub
If AudiostationMP3Player.CurrentTrackNumber = Mp3Playlist.StorageContainer.Count And Not RepeatTrack Or Not AutoNext And Not force Then
    Call StopPlay
    Exit Sub
End If

If (Not RepeatTrack And AutoNext) Or force Then
    If TrackNumber > 0 Then
        'Track number is set by parameter
        AudiostationMP3Player.CurrentTrackNumber = TrackNumber
    Else
        'Auto select track number
        AudiostationMP3Player.CurrentTrackNumber = AudiostationMP3Player.CurrentTrackNumber + 1
    End If
End If

AudiostationMP3Player.CurrentTrackNumber = CurrentTrackNumber
mediaFilename = Mp3Playlist.GetItemByIndex(CurrentTrackNumber, 1)

CurrentMediaFilename = mediaFilename

Call StartPlay
End Sub
Public Sub PreviousTrack()
Dim mediaFilename As String

If Mp3Playlist.StorageContainer.Count = 0 Or CurrentTrackNumber = 1 Then: Exit Sub

AudiostationMP3Player.CurrentTrackNumber = CurrentTrackNumber - 1

mediaFilename = Mp3Playlist.GetItemByIndex(CurrentTrackNumber, 1)
CurrentMediaFilename = mediaFilename

Call StartPlay
End Sub
