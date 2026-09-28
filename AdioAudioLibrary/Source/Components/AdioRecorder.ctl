VERSION 5.00
Begin VB.UserControl AdioRecorder 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image Image_Main 
      Height          =   480
      Left            =   0
      Picture         =   "AdioRecorder.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "AdioRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'// Public vars
Public State As enumAdioRecorderState
Public Method As enumAdioRecordingMethod

'// Const vars
Const OFS_MAXPATHNAME = 128
Const OF_CREATE = &H1000
Const OF_READ = &H0
Const OF_WRITE = &H1

'// Structures
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

'// Enums
Public Enum enumAdioRecordingMethod
    [DirectSound]
    [LoopBackRecording]
End Enum

Public Enum enumAdioRecorderState
    [Stopped]
    [Recording]
End Enum

'// Events
Public Event Error(ErrorCode As Long, ErrorDescription As String)
Public Event Recording()
Public Event RecordingSaved(File As String)

'// Private vars
Private IsDeviceSet As Boolean
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Function RuntimeInSec() As Long

End Function
Public Function SetDeviceById(Id As Long) As Boolean

End Function
Public Function SetDevice(device As mdlAdioDevice) As Boolean

End Function
Public Sub SetVolume()

End Sub
Public Sub StartRecording()
If Not IsDeviceSet Then: RaiseEvent Error(100, "Recording device not specified"): Exit Sub

If (recPtr) Then
    Call BASS_StreamFree(chan)
    Call GlobalFree(ByVal recPtr)
    
    recPtr = 0
    chan = 0
End If

' allocate initial buffer and make space for WAVE header
recPtr = GlobalAlloc(GMEM_FIXED, BUFSTEP)
reclen = 44

' fill the WAVE header
wf.wFormatTag = 1
wf.nChannels = 2
wf.wBitsPerSample = 16
wf.nSamplesPerSec = 44100
wf.nBlockAlign = wf.nChannels * wf.wBitsPerSample / 8
wf.nAvgBytesPerSec = wf.nSamplesPerSec * wf.nBlockAlign

' Set WAV "fmt " header
wf.wfBlockType = &H20746D66
wf.wfBlockSize = 16

' Set WAV "RIFF" header
wr.RIFF = &H46464952
wr.riffBlockSize = 0
wr.riffBlockType = &H45564157

' set WAV "data" header
wd.dataBlockType = &H61746164
wd.dataBlockSize = 0

' copy WAV Header to Memory
Call CopyMemory(ByVal recPtr, wr, LenB(wr))        ' "RIFF"
Call CopyMemory(ByVal recPtr + 12, wf, LenB(wf))   ' "fmt "
Call CopyMemory(ByVal recPtr + 36, wd, LenB(wd))   ' "data"

' start recording @ 44100hz 16-bit stereo
rchan = BASS_RecordStart(44100, 2, 0, AddressOf RECORDINGCALLBACK, 0)

If (rchan = 0) Then
    Call GlobalFree(ByVal recPtr)
    
    recPtr = 0
    
    Exit Sub
End If
End Sub
Public Function StopRecording() As Boolean
Call BASS_ChannelStop(rchan)

rchan = 0

' complete the WAVE header
wr.riffBlockSize = reclen - 8
wd.dataBlockSize = reclen - 44

Call CopyMemory(ByVal recPtr + 4, wr.riffBlockSize, LenB(wr.riffBlockSize))
Call CopyMemory(ByVal recPtr + 40, wd.dataBlockSize, LenB(wd.dataBlockSize))

' Create a stream from the recording
chan = BASS_StreamCreateFile(BASSTRUE, recPtr, 0, reclen, 0)

If (chan) Then: StopRecording = True
End Function
Public Function SaveRecording(File As String) As Boolean
Dim FileHandle As Long, ret As Long, OF As OFSTRUCT

FileHandle = OpenFile(File, OF, OF_CREATE)

If (FileHandle = 0) Then
    RaiseEvent Error(100, "Can't save record file: " & File)
    Exit Function
End If

Call WriteFile(FileHandle, ByVal recPtr, reclen, ret, ByVal 0&)
Call CloseHandle(FileHandle)
End Function

Private Sub UserControl_Resize()
UserControl.Height = Image_Main.Height
UserControl.Width = Image_Main.Width
End Sub

Private Sub UserControl_Terminate()
Call GlobalFree(ByVal recPtr)
Call BASS_RecordFree
Call BASS_Free
End Sub
