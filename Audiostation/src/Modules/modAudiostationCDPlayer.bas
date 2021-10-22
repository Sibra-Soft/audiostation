Attribute VB_Name = "AudiostationCDPlayer"
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Module:           AudiostationCDPlayer
' Description:      Adds CD player functionality to the Audiostation program
'
' Date Changed:     13-10-2021
' Date Created:     04-10-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public Mode As enumCDMode
Public CurrentTrackCount As Long
Public CurrentTrackNr As Long
Private Sub EndSync(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
Dim track As Long, drive As Long, tracks As Long

track = BASS_CD_StreamGetTrack(channel)
drive = HiWord(track)
tracks = BASS_CD_GetTracks(drive)

If (tracks = -1) Then Exit Sub  ' error, eg. CD removed?

track = LoWord(track) + 1 ' next track

If (track >= tracks) Then Exit Sub  ' no more tracks

Call PlayTrackFromCD(drive, track)
End Sub
Private Sub PlayTrackFromCD(drive As Long, track As Long)
If (stream(drive)) Then
    Call BASS_CD_StreamSetTrack(stream(drive), track) ' already have a stream, so just set the track
Else
    stream(drive) = BASS_CD_StreamCreate(drive, track, 0)  ' create stream
    Call BASS_ChannelSetSync(stream(drive), BASS_SYNC_END, 0, AddressOf EndSync, 0) ' set end sync
End If

Call BASS_ChannelPlay(stream(drive), BASSFALSE) ' start playing
End Sub
Public Sub StopPlay()
Call BASS_ChannelStop(curdrive)
End Sub

Public Function CheckIfCDRomDriveExists() As Boolean
Dim a As Long, n As Long
Dim cdi As BASS_CD_INFO

a = 0
While (a < MAXDRIVES And BASS_CD_GetInfo(a, cdi) <> 0)
    a = a + 1
Wend

If (a = 0) Then
    CheckIfCDRomDriveExists = False
Else
    CheckIfCDRomDriveExists = True
End If
End Function
Public Sub Play()
AudiostationMIDIPlayer.StopMidiPlayback
AudiostationCDPlayer.StopPlay

PlayStateMediaMode = CDMediaMode

If PlayState = Paused Then
    Call BASS_ChannelPlay(curdrive, 0)
Else
    Call BASS_ChannelPlay(curdrive, 1)
End If

CurrentTrackCount = BASS_CD_GetTracks(curdrive)
End Sub
Public Sub Pause()
Call BASS_ChannelPause(curdrive)
End Sub
Public Sub Forward()
Dim pos As Long
pos = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(curdrive, BASS_POS_BYTE))
Call BASS_ChannelSetPosition(chan, BASS_ChannelSeconds2Bytes(curdrive, pos - 5), BASS_POS_BYTE)
End Sub
Public Sub Rewind()
Dim pos As Long
pos = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(curdrive, BASS_POS_BYTE))
Call BASS_ChannelSetPosition(chan, BASS_ChannelSeconds2Bytes(curdrive, pos + 5), BASS_POS_BYTE)
End Sub
Public Sub OpenOrCloseDriveDoor()
If BASS_CD_DoorIsOpen(curdrive) Then
    Call BASS_CD_Door(curdrive, BASS_CD_DOOR_CLOSE)
Else
    Call BASS_CD_Door(curdrive, BASS_CD_DOOR_OPEN)
End If
End Sub
Public Sub NextTrack()
Select Case Mode
     Case LoopMode: CurrentTrackNr = CurrentTrackNr
     Case RandomMode: CurrentTrackNr = Extensions.RandomNumber(0, CInt(CurrentTrackCount))
     Case Else: CurrentTrackNr = CurrentTrackNr + 1
End Select

Call PlayTrackFromCD(curdrive, CurrentTrackNr)
End Sub
Public Sub PreviousTrack()
CurrentTrackNr = CurrentTrackNr - 1
Call PlayTrackFromCD(curdrive, CurrentTrackNr)
End Sub

