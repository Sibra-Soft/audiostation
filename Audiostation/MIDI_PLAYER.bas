Attribute VB_Name = "AudiostationMidiPlayer"
Public MidiPlaylist As New LocalStorage
Public LastMidiDirectory As String

Public CurrentPlayerProcess As Long
Public CurrentPlayerProcessToken As String

Public CurrentMidiTrackNumber As Integer
Public CurrentMidiFile As String
Public Function StartMidiPlayback()
If Form_Plugin_MID.OutputDevCombo.ListCount = 0 Then
    MsgBox "Can't play midi files without midi device driver", vbCritical
Else
    If CurrentMidiTrackNumber = 0 Then: CurrentMidiTrackNumber = 1
    
    AudiostationMP3Player.StopPlay
    Form_Plugin_MID.StopPlay
    
    CurrentMidiFile = MidiPlaylist.GetItemByIndex(CurrentMidiTrackNumber, 1)
    CurrentPlayerProcessToken = "WDfgQ8Ds0uidpobuk7l55xzM3"
    
    Form_Main.lbl_Midi_Filename.Caption = Extensions.GetFileNameFromFilePath(CurrentMidiFile, False)
    
    Select Case LCase(Right(CurrentMidiFile, 3))
        Case "mid", "kar"
            Form_Main.lbl_Midi_Filename.Caption = Extensions.GetFileNameFromFilePath(CurrentMidiFile, False)
            Call Form_Plugin_MID.StartPlay(CurrentMidiFile)
        
        Case "sid"
            Call Extensions.TerminateProcessByPid(CurrentPlayerProcess)
            CurrentPlayerProcess = Shell(App.path & "\support\sidplayer\sid_player.exe " & Chr(34) & CurrentMidiFile & Chr(34) & " -o1", vbHide)
            
        Case "mus"
            Call Extensions.TerminateProcessByPid(CurrentPlayerProcess)
            CurrentPlayerProcess = Shell(App.path & "\support\beepsymphony\beepsymphony.exe " & Chr(34) & CurrentMidiFile & Chr(34) & "#" & CurrentPlayerProcessToken, vbHide)
    End Select
    
    Form_Main.Trm_Lights_Midi.Tag = 1
    Form_Main.Trm_Lights_Midi.Enabled = True
    
    Form_Main.Trm_Midi_Play.Tag = 0
    Form_Main.Trm_Midi_Play.Enabled = True
End If
End Function
Public Function StopMidiPlayBack()
Form_Plugin_MID.StopPlay

Call Extensions.TerminateProcessByPid(CurrentPlayerProcess)

Form_Main.Trm_Lights_Midi.Tag = 2
Form_Main.Trm_Midi_Play.Enabled = False
End Function
Public Function PreviousMidiTrack()
AudiostationMidiPlayer.StartMidiPlayback
End Function
Public Function NextMidiTrack()
AudiostationMidiPlayer.StartMidiPlayback
End Function
