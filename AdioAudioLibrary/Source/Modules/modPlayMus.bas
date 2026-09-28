Attribute VB_Name = "modPlayMus"
Option Explicit

'///////////////////////////////////////////////////////////////
'// FileName        : modPlayMus.bas
'// FileType        : Microsoft Visual Basic 6 - Module
'// Author          : Alex van den Berg
'// Created         : 15-03-2026
'// Last Modified   : 18-03-2026
'// Copyright       : Sibra-Soft
'// Description     : *.mus file play module
'////////////////////////////////////////////////////////////////

Public Type Note
    Note As Long
    length As Long
    octave As Long
    Staccato As Boolean
    tie As Boolean
End Type

Private Const SAMPLE_RATE As Long = 44100

Private repPoint As Integer

Public Tempo As Long
Private mStream As Long

Private mPhase As Double

Private Notes() As Note
'*
'* Start playback
'*
Public Sub StartPlay()
Dim I As Integer

Call AudioDone
Call AudioInit

For I = LBound(Notes) To UBound(Notes)
    Note Notes(I)
    DoEvents
Next I
End Sub
'*
'* Stop playback
'*
Public Sub StopPlay()
Call AudioDone
End Sub
Private Sub AudioInit()
Call BASS_SetConfig(BASS_CONFIG_UPDATEPERIOD, 10)
Call BASS_SetConfig(BASS_CONFIG_BUFFER, 50)

If BASS_Init(-1, SAMPLE_RATE, 0, 0, 0) = 0 Then
    Exit Sub
End If

mStream = BASS_StreamCreate(SAMPLE_RATE, 1, BASS_SAMPLE_FLOAT, STREAMPROC_PUSH, 0)
If mStream = 0 Then
    Exit Sub
End If

Call BASS_ChannelPlay(mStream, 0)
End Sub
Private Sub AudioDone()
If mStream <> 0 Then
    Call BASS_ChannelStop(mStream)
    Call BASS_StreamFree(mStream)
    
    mStream = 0
End If

Call BASS_Free
End Sub
Private Sub Note(Note As Note)
Dim t As Long, tie As Integer

DoEvents

If Not Note.tie Then tie = 0 '25
If Note.length - tie < 0 Then tie = 0

If Note.Note <> 0 Then
    PlayNote Note.Note * (2 ^ Note.octave), 40
    PlaySilence Note.length * (240 / Tempo) - tie
Else
    PlaySilence Note.length * (240 / Tempo) - tie
End If
End Sub
Private Sub PlayNote(ByVal freq As Double, ByVal durationMs As Long, Optional ByVal volume As Double = 0.25)
Dim sampleCount As Long
Dim buf() As Single
Dim I As Long
Dim stepRad As Double

sampleCount = (SAMPLE_RATE * durationMs) \ 1000
If sampleCount <= 0 Then Exit Sub

ReDim buf(0 To sampleCount - 1) As Single

stepRad = 6.28318530717959 * freq / SAMPLE_RATE

For I = 0 To sampleCount - 1
    buf(I) = CSng(Sin(mPhase) * volume)
    mPhase = mPhase + stepRad
    If mPhase >= 6.28318530717959 Then
        mPhase = mPhase - 6.28318530717959
    End If
Next I

' lengte in bytes: Single = 4 bytes
Call BASS_StreamPutData(mStream, buf(0), sampleCount * 4)
End Sub
Private Sub PlaySilence(ByVal durationMs As Long)
Dim sampleCount As Long
Dim buf() As Single

sampleCount = (SAMPLE_RATE * durationMs) \ 1000
If sampleCount <= 0 Then Exit Sub

ReDim buf(0 To sampleCount - 1) As Single

Call BASS_StreamPutData(mStream, buf(0), sampleCount * 4)
End Sub
'*
'* Load new *.mus file
'* @param String strFile: The file to load
'* @returns Boolean: True if the file is loaded successfully
'*
Public Sub LoadFile(strFile As String)
Dim tmpStr As String, t As Long, Resp As Long, tBool As String, tBoolB As String
Dim asBinary As Boolean, strOct As String, tmpNoteStorage() As Integer, y As Integer
Dim tmpLenStorage() As Integer, strNot As String, strLen As String, x As Integer

Err.Clear

On Error Resume Next
If Err.number = 0 Then
    repPoint = 0
    
    Open strFile For Binary As #1
        tmpStr = Space(3)
        Get #1, , tmpStr
        asBinary = (tmpStr = "BIN")
        
        If asBinary Then
            Get #1, , t
            tmpStr = t
            Get #1, , t
            
            ReDim Notes(1 To tmpStr)
            
            tmpStr = Space(UBound(Notes) / 8)
            If UBound(Notes) / 8 > Len(tmpStr) Then tmpStr = tmpStr & " "
            Get #1, , tmpStr
            For t = 1 To Len(tmpStr)
                tBoolB = tBoolB & format(DecToBin(Asc(mId(tmpStr, t, 1))), "00000000")
            Next t

            tmpStr = Space(UBound(Notes) / 8)
            If UBound(Notes) / 8 > Len(tmpStr) Then tmpStr = tmpStr & " "
            Get #1, , tmpStr
            For t = 1 To Len(tmpStr)
                tBool = tBool & format(DecToBin(Asc(mId(tmpStr, t, 1))), "00000000")
            Next t

            tmpStr = Space(UBound(Notes) / 2)
            If UBound(Notes) / 2 > Len(tmpStr) Then tmpStr = tmpStr & " "
            Get #1, , tmpStr
            For t = 1 To Len(tmpStr)
                strOct = strOct & format(DecToBin(Asc(mId(tmpStr, t, 1))), "00000000")
            Next t

            Get #1, , t
            ReDim tmpNoteStorage(t)
            Get #1, , tmpNoteStorage
            x = Len(DecToBin(Str(UBound(tmpNoteStorage))))

            tmpStr = Space(x * UBound(Notes) / 8)
            If x * UBound(Notes) / 8 > Len(tmpStr) Then tmpStr = tmpStr & " "

            Get #1, , tmpStr
            For t = 1 To Len(tmpStr)
                strNot = strNot & format(DecToBin(Asc(mId(tmpStr, t, 1))), "00000000")
            Next t

            Get #1, , t
            ReDim tmpLenStorage(t)
            Get #1, , tmpLenStorage
            y = Len(DecToBin(Str(UBound(tmpLenStorage))))

            tmpStr = Space(y * UBound(Notes) / 8)
            If y * UBound(Notes) / 8 > Len(tmpStr) Then tmpStr = tmpStr & " "

            Get #1, , tmpStr
            For t = 1 To Len(tmpStr)
                strLen = strLen & format(DecToBin(Asc(mId(tmpStr, t, 1))), "00000000")
            Next t

            For t = 1 To UBound(Notes)
                With Notes(t)
                    .Staccato = (mId(tBoolB, t, 1) = "1")
                    .tie = (mId(tBool, t, 1) = "1")
                    .octave = BinToDec(mId(strOct, (t - 1) * 4 + 1, 4)) - 4
                    .Note = tmpNoteStorage(BinToDec(mId(strNot, (t - 1) * x + 1, x)))
                    .length = tmpLenStorage(BinToDec(mId(strLen, (t - 1) * y + 1, y)))
                End With
            Next t
        End If
    Close #1
End If
End Sub
Private Function BinToDec(bin As String) As Long
Dim t As Integer

For t = 1 To Len(bin)
    BinToDec = BinToDec * 2 + Val(mId(bin, t, 1))
Next t
End Function
Private Function DecToBin(ByVal dec As String) As String
Dim t As Integer

Do While dec > 0
     DecToBin = IIf(dec / 2 = Int(dec / 2), "0", "1") & DecToBin
     dec = Int(dec / 2)
Loop

If DecToBin = "" Then DecToBin = "0"
End Function
