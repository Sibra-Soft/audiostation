Attribute VB_Name = "modMidi"
Public Sub MidiStackPushCommon(ByRef backupelement As Integer, ByRef OutputComponent As MidiOutput)
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
If OutputComponent.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close ' must not run until verified

backupelement = OutputComponent.StackPush(0) ' zero for next available

If backupelement = 0 Then: Err.Raise 1, , "PROGRAM ERROR 3875, forgot to pop somewhere before"

Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
Public Sub MidiStackPopCommon(ByRef backupelement As Integer, ByRef OutputComponent As MidiOutput)
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
If OutputComponent.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close ' must not run until verified

If backupelement = 0 Then  ' nothing to restore
ElseIf OutputComponent.StackPop(backupelement) = backupelement Then ' okay
Else
    Err.Raise 1, , "PROGRAM ERROR 3876, something interrupted the previous push"
End If

Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
