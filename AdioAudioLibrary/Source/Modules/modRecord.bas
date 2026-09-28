Attribute VB_Name = "modRecord"
Option Explicit

' MEMORY
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2

Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' WAV Header
Public Type WAVEHEADER_RIFF    ' == 12 bytes ==
    RIFF As Long                ' "RIFF" = &H46464952
    riffBlockSize As Long       ' reclen - 8
    riffBlockType As Long       ' "WAVE" = &H45564157
End Type

Public Type WAVEFORMAT         ' == 24 bytes ==
    wfBlockType As Long         ' "fmt " = &H20746D66
    wfBlockSize As Long
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
End Type

Public Type WAVEHEADER_data    ' == 8 bytes ==
    dataBlockType As Long        ' "data" = &H61746164
    dataBlockSize As Long        ' reclen - 44
End Type

Public wr As WAVEHEADER_RIFF
Public wf As WAVEFORMAT
Public wd As WAVEHEADER_data

Public BUFSTEP As Long        ' memory allocation unit
Public input_ As Long         ' current input source
Public recPtr As Long         ' a recording pointer to a memory location
Public reclen As Long         ' buffer length

Public rchan As Long          ' recording channel
Public chan As Long           ' playback channel
Public Function RECORDINGCALLBACK(ByVal handle As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long
If ((reclen Mod BUFSTEP) + length >= BUFSTEP) Then
    recPtr = GlobalReAlloc(ByVal recPtr, ((reclen + length) / BUFSTEP + 1) * BUFSTEP, GMEM_MOVEABLE)
    If recPtr = 0 Then
        rchan = 0
        Debug.Print "Out of memory!"
        
        RECORDINGCALLBACK = BASSFALSE ' stop recording
        Exit Function
    End If
End If

Call CopyMemory(ByVal recPtr + reclen, ByVal buffer, length)

reclen = reclen + length
RECORDINGCALLBACK = BASSTRUE ' continue recording
End Function
Public Sub UpdateInputInfo()
Dim it As Long
Dim level As Single

it = BASS_RecordGetInput(input_, level) ' get info on the input
If (it = -1 Or level < 0) Then ' failed
    Call BASS_RecordGetInput(-1, level) ' try master input instead
    If (level < 0) Then level = 1 ' that failed too, just display 100%
End If
 
Debug.Print ("Level: " & level * 100)

Dim type_ As String
Select Case (it And BASS_INPUT_TYPE_MASK)
    Case BASS_INPUT_TYPE_DIGITAL:
        type_ = "digital"
    Case BASS_INPUT_TYPE_LINE:
        type_ = "line-in"
    Case BASS_INPUT_TYPE_MIC:
        type_ = "microphone"
    Case BASS_INPUT_TYPE_SYNTH:
        type_ = "midi synth"
    Case BASS_INPUT_TYPE_CD:
        type_ = "analog cd"
    Case BASS_INPUT_TYPE_PHONE:
        type_ = "telephone"
    Case BASS_INPUT_TYPE_SPEAKER:
        type_ = "pc speaker"
    Case BASS_INPUT_TYPE_WAVE:
        type_ = "wave/pcm"
    Case BASS_INPUT_TYPE_AUX:
        type_ = "aux"
    Case BASS_INPUT_TYPE_ANALOG:
        type_ = "analog"
    Case Else:
        type_ = "undefined"
End Select

Debug.Print type_
End Sub
