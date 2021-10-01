VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public FileName As String
'Play a music file
Public Sub Play()
    mciSendString "close song", 0, 0, 0
    FileName = Chr(34) + Trim(FileName) + Chr(34)
    mciSendString "open " & FileName & " alias song", 0, 0, 0
    mciSendString "play song", 0, 0, 0
End Sub
'Used to convert the short filename to the long filename
Public Function GetLongFilename(ByVal sShortName As String) As String
    Dim sLongName As String
    Dim sTemp As String
    Dim iSlashPos As Integer

    'Add \ to short name to prevent Instr from failing
    sShortName = sShortName & "\"

    'Start from 4 to ignore the "[Drive Letter]:\" characters
    iSlashPos = InStr(4, sShortName, "\")

    'Pull out each string between \ character for conversion
    Do While iSlashPos
        sTemp = Dir(Left$(sShortName, iSlashPos - 1), _
          vbNormal + vbHidden + vbSystem + vbDirectory)
        If sTemp = "" Then
            'Error 52 - Bad File Name or Number
            GetLongFilename = ""
            Exit Function
        End If
        sLongName = sLongName & "\" & sTemp
        iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
    Loop

    'Prefix with the drive letter
    GetLongFilename = Left$(sShortName, 2) & sLongName
End Function
'Check if the song has ended. Good for repeat play, ect.
Public Function EndOfSong() As Boolean
    Dim ThePos As Integer
    Dim TheDur As Integer
    ThePos = GetPositioninSec()
    TheDur = GetDurationInSec()
    If ThePos = 0 Or TheDur = 0 Then Exit Function
    If ThePos = TheDur Then
        EndOfSong = True
    Else
        EndOfSong = False
    End If
End Function
'Returns the current device in use
Public Function deviceName() As String
    Dim TheData As String
    TheData = Space$(128)
    mciSendString "sysinfo song installname", TheData, 128, 0
    deviceName = TheData
End Function
'Mute just the right channel. True = Mute,False = UnMute
Public Function SetRightChannelOff(ByVal Off As Boolean) As Boolean
    If Off = True Then
        mciSendString "setaudio song right off", 0, 0, 0
    Else
        mciSendString "setaudio song right on", 0, 0, 0
    End If
End Function
'Mute just the left channel. True = Mute,False = UnMute
Public Function SetLeftChannelOff(ByVal Off As Boolean) As Boolean
    If Off = True Then
        mciSendString "setaudio song left off", 0, 0, 0
        SetLeftChannelOff = True
    Else
        mciSendString "setaudio song left on", 0, 0, 0
        SetLeftChannelOff = False
    End If
End Function
'Get the time remaining in milli-seconds
Public Function GetTimeRemaininginMS() As Integer
    GetTimeRemaininginMS = GetDurationInMS() - GetPositioninMS()
End Function
'Get the time remaining in seconds
Public Function GetTimeRemaininginSec() As Integer
    GetTimeRemaininginSec = GetDurationInSec() - GetPositioninSec()
End Function
'Get the value of the playing speed
Public Function GetSpeed() As Integer
    Dim Rate As String
    Rate = Space$(30)
    mciSendString "status song play speed", Rate, 30, 0
    Rate = CLng(Rate)
    GetSpeed = Rate
End Function
'Get the time remaining in a user friendly format
Public Function GetFormatTimeRemaining() As String
    GetFormatTimeRemaining = GetThisTime(GetTimeRemaininginMS)
End Function
'Used to find the last backslash of the file path
Private Function GetLastBackSlash(ByVal Text As String) As String
    Dim I, pos As Integer
    Dim lastslash As Integer
    For I = 1 To Len(Text)
        pos = InStr(I, Text, "\", vbTextCompare)
        If pos <> 0 Then lastslash = pos
    Next I
    GetLastBackSlash = Right(Text, Len(Text) - lastslash)
End Function
'Used for getting the Right/Left of a string
Private Function RightLeft(ByVal Source As String, ByVal token As String) As String
    Dim I As Integer
    RightLeft = ""
    For I = Len(Source) To 1 Step -1
        If Mid(Source, I, 1) = token Then
            RightLeft = Left(Source, I - 1)
            Exit Function
        End If
    Next I
End Function
'Set the left channel to a desired volume
Public Function SetLeftVolume(ByVal Value As Integer) As Integer
    mciSendString "setaudio song left volume to " & Value, 0, 0, 0
End Function
'Set the right channel volume
Public Function SetRightVolume(ByVal Value As Integer) As Integer
    '1000 = max | 0 = min
    mciSendString "setaudio song right volume to " & Value, 0, 0, 0
End Function
'Set all audio off(mute)
Public Function setAudioOff()
    mciSendString "set song audio all off", 0, 0, 0
End Function
'Set all audio on(Unmute)
Public Function setAudioOn()
    mciSendString "set song audio all on", 0, 0, 0
End Function
'Set the speed for the song to be played
Public Function SetSpeed(ByVal speed As Integer)
    mciSendString "set song speed " & speed, "", 0, 0
End Function
'Set the volume for both channels
Public Function SetVolume(ByVal Value As Integer) As Integer
    '1000 = max | 0 = min
    mciSendString "setaudio song volume to " & Value, 0, 0, 0
End Function

'Get the position in MilliSeconds
Public Function GetPositioninMS() As Long
    Static Stat As String
    Stat = Space$(128)
    mciSendString "set song time format milliseconds", 0, 0, 0
    mciSendString "status song position", Stat, 128, 0
    GetPositioninMS = Val(Trim(Stat))
End Function
'Get the Position in Seconds
Public Function GetPositioninSec() As Integer
    GetPositioninSec = Val(GetPositioninMS() \ 1000)
End Function
'Get the positon in a user friendly format
Public Function GetFormatPosition() As String
    Dim sec As Integer
    Dim Mins As Integer
    sec = Val(GetPositioninSec())
    If sec < 60 Then GetFormatPosition = "0:" & format(sec, "00")
    If sec > 59 Then
        Mins = Int(sec / 60)
        sec = sec - (Mins * 60)
        GetFormatPosition = format(Mins, "0") & ":" & format(sec, "00")
    End If
End Function
'Calculates the bitrate of a mp3 file only
Public Function GetMP3Bitrate() As Integer
    On Error Resume Next
    Dim ThePath As String
    Dim TheSize As Integer
    Dim Ext As String
    ThePath = FileName
    Ext = Microsoft.VisualBasic.Strings.Right(ThePath, 4)
    If Ext = "mp3" & Chr(34) Then
        'Strip off the quotations from the filename that was added with Chr$(34)
        ThePath = FileName
        TheSize = Len(ThePath) - 1
        ThePath = Right(ThePath, TheSize)
        ThePath = Left(ThePath, Len(ThePath) - 1)
        '
        TheSize = FileLen(ThePath)
        TheSize = TheSize * 8

        GetMP3Bitrate = TheSize \ GetDurationInSec()
    Else
        GetMP3Bitrate = 0
    End If

End Function
'Change the position of the song in seconds
Public Sub ChangePosition(ByVal TheSecond As Long)
    TheSecond = TheSecond * 1000
    If isPlaying() = True Then mciSendString "play song from " & TheSecond, 0, 0, 0
    If isPlaying() = False Then mciSendString "seek song to " & TheSecond, 0, 0, 0
End Sub
'Close the mci device
Public Function CloseAudio()
    mciSendString "close all", 0, 0, 0
End Function
'Stop playing the song
Public Sub StopPlay()
    mciSendString "stop song", 0, 0, 0
End Sub
'Pause the song
Public Sub Pause()
    mciSendString "pause song", 0, 0, 0
End Sub
'Get duration in Milliseconds
Public Function GetDurationInMS() As Long
    Dim TotalTime As String
    TotalTime = Space$(128)
    mciSendString "status song length", TotalTime, 128, 0&
    GetDurationInMS = Val(Trim(TotalTime))
End Function
'Get the bitrate in kb instead of bits
Public Function GetFormatMP3Bitrate() As Integer
    GetFormatMP3Bitrate = GetMP3Bitrate() \ 1000
End Function
'Get the left channel volume
Public Function GetLeftVolume() As Integer
    Dim LeftLevel As String
    LeftLevel = Space(128)
    mciSendString "status song left volume", LeftLevel, 128, 0
    GetLeftVolume = Val(LeftLevel)
End Function
'Get the right channel volume value
Public Function GetRightVolume() As Integer
    Dim Rightlevel As String
    Rightlevel = Space$(128)
    mciSendString "status song right volume", Rightlevel, 128, 0
    GetRightVolume = Val(Rightlevel)
End Function
'Get the volume value for both channels
Public Function GetVolume() As Integer
    Dim TheLevel As String
    TheLevel = Space$(128)
    mciSendString "status song volume", TheLevel, 128, 0
    GetVolume = Val(TheLevel)
End Function
'Get the duration in a user friendly format
Public Function GetFormatDuration() As String
    Dim Stat As String
    Dim T As String
    Dim lTotalTime As Integer
    Stat = Space$(128)
    mciSendString "set song time format ms", Stat, 128, 0&
    mciSendString "status song length", Stat, 128, 0&
    lTotalTime = Val(Stat)
    T = GetThisTime(lTotalTime)
    GetFormatDuration = T
End Function
'Used to format the position and duration
Private Function GetThisTime(ByVal timein As Integer) As String
    Dim conH As Integer
    Dim conM As Integer
    Dim conS As Integer
    Dim remTime As Integer
    Dim strRetTime As String
    remTime = timein / 1000
    conH = Int(remTime / 3600)
    remTime = remTime Mod 3600
    conM = Int(remTime / 60)
    remTime = remTime Mod 60
    conS = remTime
    If conH > 0 Then
        strRetTime = Trim(str(conH)) & ":"
    Else
        strRetTime = ""
    End If
    If conM >= 10 Then
        strRetTime = strRetTime & Trim(str(conM))
    ElseIf conM > 0 Then
        strRetTime = strRetTime & Trim(str(conM))
    Else
        strRetTime = strRetTime & "0"
    End If
    strRetTime = strRetTime & ":"
    If conS >= 10 Then
        strRetTime = strRetTime & Trim(str(conS))
    ElseIf conS > 0 Then
        strRetTime = strRetTime & "0" & Trim(str(conS))
    Else
        strRetTime = strRetTime & "00"
    End If
    GetThisTime = strRetTime
End Function
'Get the Duration in Seconds
Public Function GetDurationInSec() As Integer
    GetDurationInSec = GetDurationInMS() \ 1000
End Function
'Resume a paused song
Public Sub ResumePlay()
    mciSendString "resume song", 0, 0, 0
End Sub
'Check the status to see if the music is playing, stopped, ect.
Public Function GetStatus() As String
    Dim Status As String
    Status = Space$(128)
    mciSendString "status song mode", Status, 128, 0
    GetStatus = Status
End Function
'Check to see if a song is playing
Public Function isPlaying() As Boolean
    Dim Stat As String
    Stat = Space$(7)
    mciSendString "status song mode", Stat, 128, 0
    If Stat = "playing" Then
        isPlaying = True
    Else
        isPlaying = False
    End If
End Function
