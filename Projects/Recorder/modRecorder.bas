Attribute VB_Name = "modRecorder"
Option Explicit

Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

' MEMORY
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' FILE
Const OFS_MAXPATHNAME = 128
Const OF_CREATE = &H1000
Const OF_READ = &H0
Const OF_WRITE = &H1

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' WAV Header
Private Type WAVEHEADER_RIFF    ' == 12 bytes ==
    wrBlockTypeRiff As Long     ' "RIFF", The characters "RIFF" indicate the start of the RIFF header
    wrBlockSize As Long         ' FileSize – 8, This is the size of the entire file following this data, i.e., the size of the rest of the file
    wrBlockTypeWave As Long     ' "WAVE", The characters "WAVE" indicate the format of the data
End Type

Private Type WAVEHEADER_data    ' == 8 bytes ==
    wdBlockTypeData As Long     ' "data", The "data" characters specify that the audio data is next in the file
    wdBlockSize As Long         ' The length of the data in bytes - WaveHeader (44)
End Type

Private Type WAVEFORMATEX       ' == 44 bytes ==
    riff As WAVEHEADER_RIFF     ' "RIFF"
    wfBlockTypeFmt As Long      ' "fmt ", The "fmt " characters specify that this is the section of the file describing the format specifically
    wfBlockSize As Long         ' 16 bytes, The size of the WAVEFORMATEX data to follow below
    wFormatTag As Integer       ' 2 bytes, Only PCM data is supported in this sample
    nChannels As Integer        ' 2 bytes, Number of channels in (1 for mono, 2 for stereo)
    nSamplesPerSec As Long      ' 4 bytes, Sample rate of the waveform in samples per second
    nAvgBytesPerSec As Long     ' 4 bytes, Average bytes per second which can be used to determine the time-wise length of the audio
    nBlockAlign As Integer      ' 2 bytes, Specifies how each audio block must be aligned in bytes
    wBitsPerSample As Integer   ' 2 bytes, How many bits represent a single sample (typically 8 or 16)
    data As WAVEHEADER_data     ' "data"
End Type

Dim wf As WAVEFORMATEX          ' wave format

Public BUFSTEP As Long          ' memory allocation unit

Public indev As Long            ' current input device
Public instream As Long         ' input stream
Public inmixer As Long          ' mixer for resampling input

Public outdev As Long           ' output device
Public outstream As Long        ' playback stream
Public outmixer As Long         ' mixer for resampling output

Public recPtr As Long           ' a recording pointer to a memory location
Public reclen As Long           ' buffer length

Public inlevel As Single        ' input level

' display error messages
Public Sub Error_(ByVal es As String)
    Call MessageBox(Form_Main.hwnd, es & vbCrLf & vbCrLf & "error code: " & BASS_ErrorGetCode, "Error", vbExclamation)
End Sub

' WASAPI input processing function
Function InWasapiProc(ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long
    Dim temp(50000) As Byte, c As Long

    ' give the data to the mixer feeder stream
    Call BASS_StreamPutData(instream, ByVal buffer, length)

    ' get back resampled data from the mixer
    Do
        c = BASS_ChannelGetData(inmixer, temp(0), UBound(temp) + 1)
        If (c > 0) Then
            ' increase buffer size if needed
            If ((reclen Mod BUFSTEP) + c >= BUFSTEP) Then
                recPtr = GlobalReAlloc(ByVal recPtr, ((reclen + c) / BUFSTEP + 1) * BUFSTEP, GMEM_MOVEABLE)
                If (recPtr = 0) Then
                    Call Error_("Out of memory!")
                    Form_Main.btnRecord.Caption = "Record"
                    InWasapiProc = 0 ' stop recording
                    Exit Function
                End If
            End If
            ' buffer the data
            Call CopyMemory(ByVal recPtr + reclen, temp(0), c)
            reclen = reclen + c
        End If
    Loop While (c > 0)
    InWasapiProc = 1 ' continue recording
End Function

' WASAPI output processing function
Function OutWasapiProc(ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long
    Dim c As Long
    c = BASS_ChannelGetData(outmixer, ByVal buffer, length)
    If (c < 0) Then ' at the end
        If (BASS_WASAPI_GetData(0, BASS_DATA_AVAILABLE) = 0) Then ' no buffered data remaining, so...
            Call BASS_WASAPI_Stop(BASSFALSE) ' stop the output
        End If
        OutWasapiProc = 0
        Exit Function
    End If
    OutWasapiProc = c
End Function

Public Sub StartRecording()
    Dim rate As Long
    If (recPtr) Then ' free old recording...
        Call BASS_StreamFree(outstream)
        outstream = 0
        Call GlobalFree(ByVal recPtr)
        recPtr = 0
        Form_Main.btnPlay.Enabled = False
        Form_Main.btnSave.Enabled = False
    End If
    
    ' get the sample rate choice
    rate = 48000

    ' allocate initial buffer and make space for WAVE header
    recPtr = GlobalAlloc(GMEM_FIXED, BUFSTEP)
    reclen = LenB(wf)   ' 44

    ' fill the WAVE header
    With wf
        .riff.wrBlockTypeRiff = &H46464952         ' "RIFF"
        .riff.wrBlockSize = 0                      ' after recording
        .riff.wrBlockTypeWave = &H45564157         ' "WAVE"

        .wfBlockTypeFmt = &H20746D66               ' "fmt "
        .wfBlockSize = 16
        .wFormatTag = 1
        .nChannels = 2
        .wBitsPerSample = 16
        .nSamplesPerSec = rate
        .nBlockAlign = .nChannels * .wBitsPerSample / 8
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign

        .data.wdBlockTypeData = &H61746164          ' "data"
        .data.wdBlockSize = 0                       ' after recording

        ' copy header to memory
        Call CopyMemory(ByVal recPtr, wf, reclen)   ' "RIFF" .. "WAVEfmt " .. "data"
    End With

    ' create a mixer and add the device's feeder stream to it
    inmixer = BASS_Mixer_StreamCreate(rate, 2, BASS_STREAM_DECODE)
    Call BASS_Mixer_StreamAddChannel(inmixer, instream, 0)

    ' start the input device
    If (BASS_WASAPI_SetDevice(indev) = 0 Or BASS_WASAPI_Start() = 0) Then
        Call Error_("Can't start recording")
        Call BASS_StreamFree(inmixer)
        inmixer = 0
        Call GlobalFree(ByVal recPtr)
        recPtr = 0
        Exit Sub
    End If

    Form_Main.btnRecord.Caption = "Stop"
    Form_Main.cmbRate.Enabled = False
End Sub

Public Sub StopRecording()
    ' stop the device and free the mixer
    Call BASS_WASAPI_SetDevice(indev)
    Call BASS_WASAPI_Stop(BASSTRUE)
    Call BASS_StreamFree(inmixer)
    inmixer = 0
    Form_Main.btnRecord.Caption = "Record"

    ' complete the WAVE header
    With wf
        .riff.wrBlockSize = reclen - 8
        .data.wdBlockSize = reclen - 44

        Call CopyMemory(ByVal recPtr + 4, .riff.wrBlockSize, LenB(.riff.wrBlockSize))
        Call CopyMemory(ByVal recPtr + 40, .data.wdBlockSize, LenB(.data.wdBlockSize))
    End With

    ' enable "save" button
    Form_Main.btnSave.Enabled = True

    ' re-enable rate selection
    Form_Main.cmbRate.Enabled = True

    If (outdev >= 0) Then
        ' create a stream from the recording
        outstream = BASS_StreamCreateFile(BASSTRUE, recPtr, 0, reclen, BASS_SAMPLE_FLOAT Or BASS_STREAM_DECODE)
        If (outstream) Then
            Call BASS_Mixer_StreamAddChannel(outmixer, outstream, 0)
            Form_Main.btnPlay.Enabled = True   ' enable "play" button
        End If
    End If
End Sub

Public Sub StartPlaying()
    Call BASS_WASAPI_SetDevice(outdev)
    Call BASS_WASAPI_Stop(BASSTRUE) ' flush the output device buffer (in case there is anything there)
    Call BASS_Mixer_ChannelSetPosition(outstream, 0, BASS_POS_BYTE) ' rewind output stream
    Call BASS_ChannelSetPosition(outmixer, 0, BASS_POS_BYTE) ' reset mixer
    Call BASS_WASAPI_Start   ' start the device
End Sub

' write the recorded data to disk
Public Sub WriteToDisk()
On Local Error Resume Next    ' if Cancel pressed...

With Form_Main.cmd
    .CancelError = True
    .flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
    .DialogTitle = "Save As..."
    .Filter = "WAV files|*.wav|All files|*.*"
    .DefaultExt = "wav"
    .ShowSave

    ' if cancel was pressed, exit sub
    If (Err.Number = 32755) Then Exit Sub

    ' create a file .WAV, directly from Memory location
    Dim FileHandle As Long, ret As Long, OF As OFSTRUCT

    FileHandle = OpenFile(.filename, OF, OF_CREATE)

    If (FileHandle = 0) Then
        Call Error_("Can't create the file")
        Exit Sub
    End If

    Call WriteFile(FileHandle, ByVal recPtr, reclen, ret, ByVal 0&)
    Call CloseHandle(FileHandle)
End With

Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecorderCommand", vbNullString)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecordingPos", 0)

End
End Sub

Public Sub InitInputDevice()
    ' inialize the input device (shared mode, 1s buffer & 100ms update period)
    If (BASS_WASAPI_Init(indev, 0, 0, 0, 1, 0.1, AddressOf InWasapiProc, 0)) Then
        ' create a BASS push stream of same format to feed the mixer/resampler
        Dim wi As BASS_WASAPI_INFO
        Call BASS_WASAPI_GetInfo(wi)
        instream = BASS_StreamCreate(wi.freq, wi.chans, BASS_SAMPLE_FLOAT Or BASS_STREAM_DECODE, STREAMPROC_PUSH, 0)
        If (inmixer) Then ' already recording, start the new device...
            Call BASS_Mixer_StreamAddChannel(inmixer, instream, 0)
            Call BASS_WASAPI_Start
        End If
        ' update level slider
        Dim level As Single
        level = BASS_WASAPI_GetVolume(BASS_WASAPI_CURVE_WINDOWS)
        If (level < 0) Then ' failed to get level
            level = 1 ' just display 100%
            Form_Main.sldInputLevel.Enabled = False
        Else
            Form_Main.sldInputLevel.Enabled = True
        End If
        Form_Main.sldInputLevel.value = level * 100
    Else ' failed, just set level slider to 0
        Form_Main.sldInputLevel.Enabled = False
        Form_Main.sldInputLevel.value = 0
    End If
    ' update device type display
    Dim type_ As String, di As BASS_WASAPI_DEVICEINFO
    Call BASS_WASAPI_GetDeviceInfo(indev, di)

    Select Case (di.type)
        Case BASS_WASAPI_TYPE_NETWORKDEVICE:
            type_ = "Remote Network Device"
        Case BASS_WASAPI_TYPE_SPEAKERS:
            type_ = "Speakers"
        Case BASS_WASAPI_TYPE_LINELEVEL:
            type_ = "Line In"
        Case BASS_WASAPI_TYPE_HEADPHONES:
            type_ = "Headphones"
        Case BASS_WASAPI_TYPE_MICROPHONE:
            type_ = "Microphone"
        Case BASS_WASAPI_TYPE_HEADSET:
            type_ = "Headset"
        Case BASS_WASAPI_TYPE_HANDSET:
            type_ = "Handset"
        Case BASS_WASAPI_TYPE_DIGITAL:
            type_ = "Digital"
        Case BASS_WASAPI_TYPE_SPDIF:
            type_ = "SPDIF"
        Case BASS_WASAPI_TYPE_HDMI:
            type_ = "HDMI"
        Case Else:
            type_ = "undefined"
    End Select
    If (di.flags And BASS_DEVICE_LOOPBACK) = BASS_DEVICE_LOOPBACK Then
        type_ = type_ & " (loopback)"
    End If
    Form_Main.lblInputType.Caption = "type: " & type_  ' display the type
End Sub

