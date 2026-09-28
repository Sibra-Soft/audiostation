VERSION 5.00
Begin VB.UserControl AdioMediaPlayer 
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2340
   ScaleWidth      =   3060
   Begin VB.Timer Timer_Playing 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   675
      Top             =   1350
   End
   Begin VB.Timer Timer_Stream 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   135
      Top             =   1350
   End
   Begin VB.Image Image_Main 
      Height          =   480
      Left            =   0
      Picture         =   "AdioMediaPlayer.ctx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label_StreamTitle 
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   945
      Visible         =   0   'False
      Width           =   1230
   End
End
Attribute VB_Name = "AdioMediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'///////////////////////////////////////////////////////////////
'// FileName        : AdioMediaPlayer.ctl
'// FileType        : Microsoft Visual Basic 6 - Usercontrol
'// Author          : Alex van den Berg
'// Created         : 28-10-2023
'// Last Modified   : 16-03-2026
'// Copyright       : Sibra-Soft
'// Description     : Usercontrol for audio playback
'///////////////////////////////////////////////////////////////

'// Private variables
Private MediaChannel As Long
Private StreamEnded As Boolean

'// Public variables
Public LoadedFile As String
Public LoadedFileType As enumAdioSupportedFileTypes
Public State As enumAdioPlayState
Public RepeatMode As enumAdioRepeatMode

'// Events
Public Event Paused()
Public Event Stopped()
Public Event Playing()
Public Event StartPlay()
Public Event MediaEnded()
Public Event NewMediaFile(file As String)
Public Event NewStream()
Public Event Error(Description As String, Code As Long)
Public Event Fading(Progress As Integer)
Public Event StreamBuffering(Percent As Integer)
Public Event StreamTitleChange(Title As String)
'*
'* Getter function for the current channel
'* @returns Long: The long value of the channel
'*
Public Function Channel() As Long
Channel = MediaChannel
End Function
'*
'* Set the balance of the speaker audio
'* @param Integer Value: Balance value between -1000 and 1000
'*
Public Sub SetBalance(Value As Integer)
Call modAdio.SetBalance(MediaChannel, Value)
End Sub
'*
'* Set the playback device based on a id
'* @param Long id: Id of the device to set as the current playback device
'*
Public Function SetDeviceById(id As Long) As Boolean

End Function
'*
'* Set the playback device based on a AdioDevice model
'* @param mdlAdioDevice device: The model of the device to set as the current playback device
'*
Public Function SetDevice(device As mdlAdioDevice) As Boolean

End Function
'*
'* Load a stream based on the specified url
'* @param String strStreamUrl: The url of the stream to load
'* @param String strProxyAddress: Proxy address to use for the stream
'*
Public Function LoadStream(strStreamUrl As String, Optional strProxyAddress As String) As Boolean
If OpenStreamByUrl(strStreamUrl) Then
    Timer_Stream.Enabled = True
    
    State = AdioPlaying
    
    RaiseEvent NewStream
End If
End Function
'*
'* Fade the current playback
'* @param enumAdioFadeType fadeType: The type to use for the fade
'* @param Integer duration: The duration of the fade
'*
Public Sub Fade(fadeType As enumAdioFadeType, Optional duration As Integer = 5)
modAdio.AdioFade MediaChannel, fadeType, duration
End Sub
'*
'* Set the volume of the mediaplayer
'* @param Integer value: 0 to 100 of the current volume
'*
Public Sub SetVolume(Value As Integer)
Call modAdio.SetVolume(MediaChannel, Value)
End Sub
'*
'* Get the current volume amount
'* @returns Integer: 0 to 100 of the current volume
'*
Public Function GetVolume() As Integer
GetVolume = modAdio.GetVolume(MediaChannel)
End Function
'*
'* Mute the current volume of the mediaplayer
'* @returns Boolean: True of False
'*
Public Function MuteAudio() As Boolean
If modAdio.Mute Then
    Call modAdio.AdioMuteOff(MediaChannel)
    Mute = False
Else
    Call modAdio.AdioMuteOn(MediaChannel)
    Mute = True
End If
End Function
'*
'* Seek the current media forward or backwards
'* @param enumAdioSeekDirection direction: The direction of the seek (forward or backward)
'* @param Integer seconds: The amount of seconds to seek
'*
Public Sub SeekBySeconds(Direction As enumAdioSeekDirection, Optional seconds As Integer = 10)
Call modAdio.AdioSeekBySeconds(MediaChannel, Direction, seconds)
End Sub
'*
'* Start playback
'*
Public Sub StartPlay()
Call modAdio.AdioPlay(MediaChannel)

StreamEnded = False
modAdio.State = AdioPlaying
State = AdioPlaying

RaiseEvent StartPlay
RaiseEvent Playing

Timer_Playing.Enabled = True
End Sub
'*
'* Stop playback
'*
Public Sub StopPlay()
If Not State = AdioPlaying Then: Exit Sub

Call modAdio.AdioStop(MediaChannel)

modAdio.State = AdioStopped
State = AdioStopped

RaiseEvent Stopped

Timer_Stream.Enabled = False
Timer_Playing.Enabled = False
End Sub
'*
'* Pause playback
'*
Public Sub PausePlay()
Call modAdio.AdioPause(MediaChannel)

modAdio.State = AdioPaused
State = AdioPaused

RaiseEvent Paused

Timer_Playing.Enabled = False
End Sub
'*
'* Get properties of the current media file
'* @returns mdlAdioProperties: Model containing the properties
'*
Public Function GetProperties() As mdlAdioProperties
Set GetProperties = modAdio.GetProperties(MediaChannel)
End Function

'*
'* Load a new media file
'* @param String strFile: The file to load
'* @returns Boolean: True if the file is loaded successfully
'*
Public Function LoadFile(strFile As String) As Boolean
Dim Fso As New FileSystemObject

If Not Helpers.FileExists(strFile) Then: RaiseEvent Error("File not found", 100)
If Not CheckFileSupport(strFile) Then: RaiseEvent Error("File not supported", 110)

Call BASS_ChannelFree(MediaChannel)

' Check the extension
Select Case Fso.GetExtensionName(strFile)

    Case "flac"
        LoadedFileType = [FLAC - Free Lossless Audio Codec File]
        MediaChannel = BASS_FLAC_StreamCreateFile(0&, StrPtr(strFile), 0&, 0&, BASS_SAMPLE_FX)
    
    Case "wma"
        LoadedFileType = [WMA - Windows Media Audio]
        MediaChannel = BASS_WMA_StreamCreateFile(0&, StrPtr(strFile), 0&, 0&, BASS_SAMPLE_FX)
    
    Case Else
        MediaChannel = BASS_StreamCreateFile(0&, StrPtr(strFile), 0&, 0&, BASS_SAMPLE_FX)

End Select

If MediaChannel Then
    State = AdioReady
    LoadedFile = strFile
    
    RaiseEvent NewMediaFile(strFile)
Else
    RaiseEvent Error("Problem while loading file: " & strFile, BASS_ErrorGetCode)
End If
End Function
Private Sub Label_StreamTitle_Change()
RaiseEvent StreamTitleChange(Label_StreamTitle.Caption)
End Sub
Private Sub Timer_Playing_Timer()
If GetProperties.RemainingInSeconds <= 0 Then: StreamEnded = True

If StreamEnded = True Then
    State = AdioEnded
    Timer_Playing.Enabled = False
    
    RaiseEvent MediaEnded
Else
    RaiseEvent Playing
End If
End Sub
Private Sub Timer_Stream_Timer()
Call TimerProc

If StreamState = Buffering Then: RaiseEvent StreamBuffering(StreamBufferProgress)

Label_StreamTitle.Caption = modAdioNetRadio.StreamMeta
End Sub
Private Sub UserControl_Resize()
width = Image_Main.width
height = Image_Main.height
End Sub
