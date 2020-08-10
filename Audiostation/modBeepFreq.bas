Attribute VB_Name = "modBeepFreq"
Public Type note
    note As Long
    Length As Long
    octave As Long
    Staccato As Boolean
    tie As Boolean
End Type

Public Declare Function Beep Lib "Kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function GetTickCount Lib "Kernel32" () As Long

Global Const ALow = 440
Global Const AshLow = 466.16
Global Const BLow = 493.88
Global Const C = 523.25
Global Const Csh = 554.37
Global Const D = 587.33
Global Const Dsh = 622.25
Global Const E = 659.26
Global Const F = 698.46
Global Const Fsh = 739.99
Global Const G = 783.99
Global Const Gsh = 830.61
Global Const a = 880
Global Const Ash = 932.33
Global Const B = 987.77
Global Const Rest = 0

Global Const Whole = 1000
Global Const Half = 0.5 * Whole
Global Const Quarter = 0.5 * Half
Global Const Eighth = 0.5 * Quarter
Global Const Sixteenth = 0.5 * Eighth
Global Const Thirtysecond = 0.5 * Sixteenth
Global Const Dotted = 1.5

Global Tempo As Long
Public Sub note(note As note)
    Dim t As Long, tie As Integer
    DoEvents
    If note.note = -1 Then
        If note.Length >= Form1.udTempo.Min And note.Length <= Form1.udTempo.Max Then
            Form1.udTempo.Value = note.Length
        End If
        Exit Sub
    End If
    If Not note.tie Then tie = 0 '25
    If note.Length - tie < 0 Then tie = 0
    If note.note <> 0 Then
        If Not note.Staccato Then
            t = note.Length * (240 / Tempo) - tie
            If t < 1 Then t = 1
            Beep note.note * (2 ^ note.octave), t
        Else
            Beep note.note * (2 ^ note.octave), 40
            t = GetTickCount
            Do While GetTickCount < (t + note.Length * (240 / Tempo)) - 40 - tie
                DoEvents
            Loop
        End If
    Else
        t = GetTickCount
        Do While GetTickCount < t + note.Length * (240 / Tempo) - tie
            DoEvents
        Loop
    End If
    
    t = GetTickCount
    Do While GetTickCount < t + tie
        DoEvents
    Loop
End Sub
