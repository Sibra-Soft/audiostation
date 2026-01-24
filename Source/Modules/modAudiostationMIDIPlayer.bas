Attribute VB_Name = "AudiostationMIDIPlayer"
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Module:           AudiostationMidiPlayer
' Description:      Adds Midi player functionality
'
' Date Changed:     25-10-2021
' Date Created:     04-10-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long

Private Const WM_CHAR = &H102
Private Const WM_CLOSE = &H10

Public MidiFilename As String
Public MidiTrackNr As Integer
Public MidiPlaystate As enumPlayStates
Public MidiPlaylist As New LocalStorage

Public ConsoleWindow As Long
Public Sub StartMidiPlayback()
Call AudiostationCDPlayer.StopPlay
Call AudiostationMP3Player.StopPlay

If PlayStateMediaMode = SidMediaMode Then Call AudiostationMIDIPlayer.StopMidiPlayback

Call Form_Main.ResetMidiVU ' Set all vu meters to 0
If MidiTrackNr = 0 Then MidiTrackNr = 1

If MidiPlaystate = Paused Then
    Call Form_Midi.StartPlay
Else
    Dim FileName As String
    
    FileName = MidiPlaylist.GetItemByIndex(MidiTrackNr, 1)
    MidiFilename = Extensions.GetFileNameFromFilePath(FileName, False)
    
    Select Case Right(FileName, 3)
        Case "sid"
            Call AudiostationMIDIPlayer.StopMidiPlayback
            PlayStateMediaMode = SidMediaMode
           
            Shell App.path & "\support\sidplayer\sid_player.exe " & Chr(34) & FileName & Chr(34), vbHide
            
            Call Extensions.Pause(500)
            ConsoleWindow = FindWindow(vbNullString, App.path & "\support\sidplayer\sid_player.exe")
    
        Case "mus"
            PlayStateMediaMode = MusMediaMode
            MsgBox "Start beep player"
            
        Case Else
            PlayStateMediaMode = MidiMediaMode
            
            Call Form_Midi.OpenFile(FileName)
            Call Form_Midi.StartPlay
            
    End Select
End If

MidiPlaystate = Playing
End Sub
Public Sub StopMidiPlayback()
If PlayStateMediaMode = SidMediaMode Then
    Call PostMessage(ConsoleWindow, WM_CLOSE, 0&, 0&)
Else
    Form_Midi.StopPlay
    MidiPlaystate = Stopped
End If
End Sub
Public Sub PauseMidiPlayback()
Form_Midi.PausePlay
MidiPlaystate = Paused
End Sub
Public Sub NextMidiTrack()
If MidiTrackNr = MidiPlaylist.StorageContainer.count Then: Exit Sub

MidiTrackNr = MidiTrackNr + 1
StartMidiPlayback
End Sub
Public Sub PreviousMidiTrack()
MidiTrackNr = MidiTrackNr - 1

If MidiTrackNr = 0 Then: Exit Sub

StartMidiPlayback
End Sub
Public Sub ForwardMidi10Seconds()
On Error Resume Next
Form_Midi.HScrollPlayerTime.value = Form_Midi.HScrollPlayerTime.value + 10
End Sub
Public Sub RewindMidi10Seconds()
On Error Resume Next
Form_Midi.HScrollPlayerTime.value = Form_Midi.HScrollPlayerTime.value - 10
End Sub
