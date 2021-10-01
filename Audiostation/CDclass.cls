VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlayerForCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long 'Get the error message of the mcidevice if any
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Dim Error As Long
Dim TheStat As String * 128
'Set the left channel to a desired volume
Public Function SetLeftVolume(ByVal Value As Integer) As Integer
    mciSendString "setaudio song left volume to " & Value, 0, 0, 0
End Function
'Set the right channel volume
Public Function SetRightVolume(ByVal Value As Integer) As Integer
    '1000 = max | 0 = min
    mciSendString "setaudio song right volume to " & Value, 0, 0, 0
End Function
'check the mci device for to see if a error occured
'and give description.
Public Function checkError() As String
    checkError = Space$(255)
    mciGetErrorString Error, checkError, Len(checkError)
End Function
'check to see if a cd is in the drive
Public Function checkForCD() As Boolean
    Error = mciSendString("status cd media present", TheStat, 128, 0)
    checkForCD = CBool(TheStat)
End Function
'determine if a drive is a cd drive
Public Function checkCDDrive(drive As String) As Boolean
    Dim cdDrive As Long
    'if return value is 5 then it is a cd-rom drive
    cdDrive = GetDriveType(drive)
    If cdDrive = 5 Then
        checkCDDrive = True
    Else
        checkCDDrive = False
    End If
End Function
'start the cdaudio mci device.
'need to call this function first
Public Function startCD(cdDrive As String)
    If checkCDDrive(cdDrive) = False Then
        MsgBox "De geselecteerde CD-Rom speler is geen CD-Rom speler", vbCritical, "Fout"
        Form_Main.Trm_CD_Play.Enabled = False
        Exit Function
    End If
    cdDrive = Chr$(34) & cdDrive & Chr$(34)
    Error = mciSendString("open " & cdDrive & " type cdaudio alias cd wait shareable", 0, 0, 0)
End Function
'start playing the cd
Public Function playCD()
    Error = mciSendString("play cd", 0, 0, 0)
End Function
'seek to the beginning of the cd
Public Function seekToBeginning()
    Error = mciSendString("seek cd to start", 0, 0, 0)
End Function
'seek to the end of the cd
Public Function seekToEnd()
    Error = mciSendString("seek cd to end", 0, 0, 0)
End Function
'stop the cd from playing
Public Function stopCD()
    Error = mciSendString("stop cd", 0, 0, 0)
End Function
'close all alias's from the mci device
Public Function closeAll()
    Error = mciSendString("close all", 0, 0, 0)
End Function
'get the current selected track
Public Function getCurrentTrack() As Long
    'getCurrentTrack = Val(Left(getPositionTMSF, 2))
    Error = mciSendString("status cd current track", TheStat, 128, 0)
    getCurrentTrack = Val(TheStat)
End Function
'get the position in seconds
Public Function getPositionSeconds() As Long
    getPositionSeconds = Val(getPositionMS) / 1000
End Function
'get the position in frames
Public Function getPositionFrames() As Long
    Dim posF As String
    Dim posMSF As String
    posMSF = Left(getPositionMSF, 8)
    posF = Right(posMSF, 2)
    getPositionFrames = Val(posF)
End Function
'get the position in minutes, seconds, and frames
Public Function getPositionMSF() As String
    Error = mciSendString("set cd time format msf", 0, 0, 0)
    Error = mciSendString("status cd position", TheStat, 128, 0)
    getPositionMSF = TheStat
End Function
'get the position in milli-seconds
Public Function getPositionMS() As Long
    Error = mciSendString("set cd time format ms", 0, 0, 0)
    Error = mciSendString("status cd position", TheStat, 128, 0)
    getPositionMS = Val(TheStat)
End Function
'get the position in tracks, minutes, seconds, and frames
Public Function getPositionTMSF() As String
    Error = mciSendString("set cd time format tmsf", 0, 0, 0)
    Error = mciSendString("status cd position", TheStat, 128, 0)
    getPositionTMSF = TheStat
End Function
'get the current time format
Public Function getTimeFormat() As String
    Error = mciSendString("status cd time format", TheStat, 128, 0)
    getTimeFormat = TheStat
End Function
'open the cd tray
Public Function setDoorOpen()
    Error = mciSendString("set cd door open", 0, 0, 0)
End Function
'close the cd tray
Public Function setDoorClosed()
    Error = mciSendString("set cd door closed", 0, 0, 0)
End Function
'get the total number of tracks on the cd
Public Function getNumberTracks() As Long
    On Error Resume Next
    Error = mciSendString("status cd number of tracks", TheStat, 128, 0)
    getNumberTracks = CLng(Trim$(TheStat))
End Function
'set the cd to play a specified track
Public Function setTrack(Track As Long)
    Error = mciSendString("play cd from " & Track, 0, 0, 0)
End Function
'pause the current playing track
Public Function pauseCD()
    Error = mciSendString("pause cd", 0, 0, 0)
End Function
'resume a paused track
Public Function resumeCD()
    Error = mciSendString("play cd", 0, 0, 0)
End Function
'fast forward a specified amount of seconds
Public Function fastForward(seconds As Long)
    Dim pos As String * 128
    Error = mciSendString("set cd time format ms", 0, 0, 0)
    Error = mciSendString("status cd position", pos, 128, 0)
    pos = CLng(Val(pos))
    If isPlaying = True Then
        Error = mciSendString("play cd from " & pos + seconds * 1000, 0, 0, 0)
    Else
        Error = mciSendString("seek cd to " & pos + seconds * 1000, 0, 0, 0)
    End If
End Function
'rewind a specified amount of seconds
Public Function fastRewind(seconds As Long)
    Dim pos As String * 128
    Error = mciSendString("set cd time format ms", 0, 0, 0)
    Error = mciSendString("status cd position", pos, 128, 0)
    pos = CLng(Val(pos))
    If isPlaying = True Then
        Error = mciSendString("play cd from " & pos - seconds * 1000, 0, 0, 0)
    Else
        Error = mciSendString("seek cd to " & pos - seconds * 1000, 0, 0, 0)
    End If
End Function
'check to see if the cd alias is playing
Public Function isPlaying() As Boolean
    Error = mciSendString("status cd mode", TheStat, 128, 0)
    If Left(TheStat, 7) = "playing" Then
        isPlaying = True
    Else
        isPlaying = False
    End If
End Function
'get the length of a track in milli-seconds
Public Function trackLengthMS(Track As Long) As Long
    Error = mciSendString("set cd time format ms", 0, 0, 0)
    Error = mciSendString("status cd length track " & Track, TheStat, 128, 0)
    trackLengthMS = Val(TheStat)
End Function
'get the length of the cd in milli-seconds
Public Function cdLengthMS() As Long
    Error = mciSendString("set cd time format ms", 0, 0, 0)
    Error = mciSendString("status cd length", TheStat, 128, 0)
    cdLengthMS = Val(TheStat)
End Function
'cd length in minute, second, frame
Public Function cdLengthMSF() As String
    Error = mciSendString("set cd time format msf", 0, 0, 0)
    Error = mciSendString("status cd length", TheStat, 128, 0)
    cdLengthMSF = TheStat
End Function
'get the length of the cd in seconds
Public Function cdLengthSeconds() As Long
    cdLengthSeconds = cdLengthMS / 1000
End Function
'track length in minute, second, frame
Public Function trackLengthMSF(Track As Long) As String
    Error = mciSendString("set cd time format msf", 0, 0, 0)
    Error = mciSendString("status cd length track " & Track, TheStat, 128, 0)
    trackLengthMSF = TheStat
End Function
'get the length of the track in seconds
Public Function trackLengthSeconds(Track As Long) As Long
    trackLengthSeconds = trackLengthMS(Track) / 1000
End Function
'set off the audio output(mute)
Public Function setAudioOff()
    Error = mciSendString("set cd audio all off", 0, 0, 0)
End Function
'set the audio on
Public Function setAudioOn()
    Error = mciSendString("set cd audio all on", 0, 0, 0)
End Function
'select the next available track
Public Function NextTrack()
    Dim Track As Long
    Track = getCurrentTrack + 1
    If Track > getNumberTracks Then Exit Function
    Error = mciSendString("play cd from " & Track, 0, 0, 0)
End Function
'select the previous available track
Public Function prevTrack()
    Dim Track As Long
    If getCurrentTrack <= 1 Then Exit Function
    Track = getCurrentTrack - 1
    If Track > getNumberTracks Then Exit Function
    Error = mciSendString("play cd from " & Track, 0, 0, 0)
End Function
