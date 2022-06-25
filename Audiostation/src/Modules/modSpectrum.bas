Attribute VB_Name = "ModBassSpectrum"
Option Explicit

Public chan As Long
Public Function Sqrt(ByVal num As Double) As Double
On Local Error GoTo isbad
Sqrt = num ^ 0.5
Exit Function
isbad:
Sqrt = 0
Err.Clear
End Function

Function Log10(ByVal X As Double) As Double
    Log10 = Log(X) / Log(10#)
End Function
Public Sub UpdateSpectrum()
Dim X As Long, Y As Long, y1 As Long
Dim fft(1024) As Single

Call BASS_ChannelGetData(chan, fft(0), BASS_DATA_FFT2048)
Call ResetSpectrum

Dim b0 As Long, BANDS As Integer
b0 = 0
BANDS = 28
Dim sc As Long, b1 As Long
Dim sum As Single
For X = 0 To BANDS - 1
    sum = 0
    b1 = 2 ^ (X * 10# / (BANDS - 1))
    If (b1 > 1023) Then b1 = 1023
    If (b1 <= b0) Then b1 = b0 + 1 ' make sure it uses at least 1 FFT bin
    sc = 10 + b1 - b0
    Do
        sum = sum + fft(1 + b0)
        b0 = b0 + 1
        
        ' Countdown from 13 to make sure we start at the bottom of the spectrum
        Call SetSpectrumBar(13 - Round((sum * 14), 0), X)
    Loop While b0 < b1
Next X
End Sub
Public Sub ResetSpectrum()
Dim I, C As Integer
For I = 0 To Form_Main.VU_Spectrum.RowCount - 1
    For C = 0 To Form_Main.VU_Spectrum.ColCount - 1
        Call Form_Main.VU_Spectrum.SetIndicatorActive(I, C, False)
    Next
Next
End Sub
Private Sub SetSpectrumBar(Row As Long, Col As Long)
Dim I As Integer

For I = Row To 13
    Call Form_Main.VU_Spectrum.SetIndicatorActive(I, Col, True)
Next
End Sub
