Attribute VB_Name = "ModSpectrum"
Option Explicit

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0&    ' color table in RGBs

Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(255) As RGBQUAD
End Type

Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal length As Long, ByVal Fill As Byte)
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

Public Const SPECWIDTH As Long = 368  ' display width
Public Const SPECHEIGHT As Long = 127 ' height (changing requires palette adjustments too)

Public chan As Long         ' stream/music handle

Public specmode As Long, specpos As Long  ' spectrum mode (and marker pos for 2nd mode)
Public specbuf() As Byte    ' a pointer

Public bh As BITMAPINFO     ' bitmap header
' MATH Functions
Public Function Sqrt(ByVal num As Double) As Double
On Local Error GoTo isbad
Sqrt = num ^ 0.5
Exit Function
isbad:
Sqrt = 0
err.Clear
End Function

Function Log10(ByVal X As Double) As Double
    Log10 = Log(X) / Log(10#)
End Function
Public Function UpdateSpectrum()
Dim X As Long, Y As Long, y1 As Long
Dim fft(1024) As Single     ' get the FFT data

Call BASS_WASAPI_GetData(fft(0), BASS_DATA_FFT2048)

ReDim specbuf(SPECWIDTH * (SPECHEIGHT + 1)) As Byte ' clear display
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
        
        Form_Main.VU_Spectrum(X).Position = sum * 100
    Loop While b0 < b1
Next X
End Function
