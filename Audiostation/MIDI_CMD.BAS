Attribute VB_Name = "MIDI_CMD"
Option Explicit

Sub all_sounds_off()
Dim channel As Integer

    For channel = 0 To 15
        Call midi_outshort(&HB0 + channel, &H78, 0)
        Call midi_outshort(&HB0 + channel, &H7B, 0)
    Next
End Sub

Sub Nibblize(ByVal Value As Long, msb As Integer, lsb As Integer)
    lsb = Value And &H7F ' screen off first seven bits
    Value = Value And &H7F80 ' take the rest
    msb = Value / 128 ' - and shift 7 bits down
End Sub

Sub all_notes_off(ch As Integer)
    Call midi_outshort(&HB0 + ch, &H78, 0)
End Sub
Sub Bendersub(ByVal ch As Integer, ByVal Value As Long)
Dim msb As Integer ' most significant byte
Dim lsb As Integer ' least significant byte

    Call Nibblize(Value, msb, lsb)

    Call midi_outshort(&HE0 + ch, lsb, msb)
End Sub

Sub control_change(ch As Integer, ccnr As Integer, ByVal v As Integer)
    Call midi_outshort(&HB0 + ch, ccnr, v)
End Sub

Sub midisetrpn(ch As Integer, pmsb As Integer, plsb As Integer, msb As Integer, lsb As Integer)
    Call midi_outshort(ch, &H65, pmsb)
    Call midi_outshort(ch, &H64, plsb)
    Call midi_outshort(ch, &H6, msb)
    Call midi_outshort(ch, &H26, lsb)
End Sub

Sub note_off(ch As Integer, ByVal kk As Integer)
    Call midi_outshort(&H90 + ch, kk, 0)
End Sub

Sub note_on(ch As Integer, ByVal kk As Integer, v As Integer)
    Call midi_outshort(&H90 + ch, kk, v)
End Sub

Sub pitch_bend(ch As Integer, l As Integer, M As Integer)
    Call midi_outshort(&HE0 + ch, l, M)
End Sub

Sub program_change(ch As Integer, cc0nr As Integer, ByVal pnr As Integer)
    Call control_change(ch, 0, cc0nr)
    Call midi_outshort(&HC0 + ch, pnr, 0)
End Sub

Sub reset_controllers()
Dim I As Integer

    For I = 0 To 15
        Call midi_outshort(&HB0 + I, &H79, 0)
    Next
End Sub

