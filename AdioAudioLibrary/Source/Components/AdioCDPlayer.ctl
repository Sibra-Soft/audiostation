VERSION 5.00
Begin VB.UserControl AdioCDPlayer 
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1905
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1410
   ScaleWidth      =   1905
   Begin VB.Image Image_Main 
      Height          =   480
      Left            =   0
      Picture         =   "AdioCDPlayer.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "AdioCDPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////////
'// FileName        : AdioCDPlayer.ctl
'// FileType        : Microsoft Visual Basic 6 - Usercontrol
'// Author          : Alex van den Berg
'// Created         : 16-08-2023
'// Last Modified   : 30-01-2026
'// Copyright       : Sibra-Soft
'// Description     : Usercontrol for CD player functionality
'////////////////////////////////////////////////////////////////

Option Explicit

'// Enums
Public Enum enumAdioCdRepeatMode
    [AdioNoRepeat]
    [AdioRandom]
    [AdioLoopTrack]
End Enum

'// Private vars
Private CDChannel As Long
Private CurDrive As Integer

'// Public vars
Public Ready As Boolean
Public CurTrack As Long
Public State As enumAdioPlayState
Public RepeatMode As enumAdioCdRepeatMode

'// Events
Public Event NoCdRomDeviceFound()
Public Event DeviceFound(id As Long, DriveName As String, DriveLetter As String)
Public Event StartPlay()
Public Event Playing()
Public Event Stopped()
Public Event Paused()
Public Event DoorChanged()
Public Event Error(ErrDescription As String, ErrCode As Long)
Public Function Channel() As Long
Channel = CDChannel
End Function
'*
'* Set the balance of the speaker audio
'* @param Integer Value: Balance value between -1000 and 1000
'*
Public Sub SetBalance(Value As Integer)
Call modAdio.SetBalance(CDChannel, Value)
End Sub
'*
'* Init function of the control
'*
Public Sub Initialize()
If GetListOfDrives.Count = 0 Then: RaiseEvent NoCdRomDeviceFound ' Check if the computer has a CD-Rom drive
End Sub
'*
'* Mute the CD player
'* @return Boolean: Tells if the audio has been muted
'*
Public Function Mute() As Boolean
If modAdio.Mute Then
    Call modAdio.AdioMuteOff(CDChannel)
    Mute = False
Else
    Call modAdio.AdioMuteOn(CDChannel)
    Mute = True
End If
End Function
'*
'* Gets player properties (songlength, duration, etc.)
'* @return AdioProperties: The properties of the current player
'*
Public Function GetProperties() As mdlAdioProperties
Dim Properties As New mdlAdioProperties

' Only return when playing
If State = AdioPlaying Then: Set Properties = modAdio.GetProperties(CDChannel)

Set GetProperties = Properties
End Function
'*
'* Seek the player forward or backwards a specified number of seconds
'* @param AdioSeekDirection Direction: The direction you want to seek (rewind, forward)
'* @param Integer Seconds: The amount of seconds you want to seek (default = 10)
'*
Public Sub SeekBySeconds(Direction As enumAdioSeekDirection, Optional Seconds As Integer = 10)
Call modAdio.AdioSeekBySeconds(CDChannel, Direction, Seconds)
End Sub
'*
'* Set the current track of the CD player
'* @return Boolean: Tells if the track has been set
'*
Public Function SetTrack(Optional TrackNr As Long = 0) As Boolean
CDChannel = BASS_CD_StreamCreate(CurDrive, TrackNr, 0&)

If CDChannel Then
    SetTrack = True
    CurTrack = TrackNr
Else
    SetTrack = False
    RaiseEvent Error("Error loading specified track", BASS_ErrorGetCode)
End If
End Function
'*
'* Pause the player
'*
Public Sub PausePlay()
Call modAdio.AdioPause(CDChannel)
State = AdioPaused

RaiseEvent Paused
End Sub
'*
'* Stop the player
'*
Public Sub StopPlay()
If Not State = AdioPlaying Then: Exit Sub

Call modAdio.AdioStop(CDChannel)
State = AdioStopped

RaiseEvent Stopped
End Sub
'*
'* Start the player
'*
Public Sub StartPlay()
Call modAdio.AdioPlay(CDChannel)

State = AdioPlaying

RaiseEvent StartPlay
RaiseEvent Playing
End Sub
'*
'* Set the current drive using a drive letter
'* @param String Letter: The letter of the drive to set as current
'* @return Boolean: Tells if the current drive has been set
'*
Public Function SetDriveByLetter(letter As String) As Boolean

End Function
'*
'* Set the current drive using the id of the drive
'* @param Integer Id: The id of the drive to set as current
'* @return Boolean: Tells if the current drive has been set
'*
Public Function SetDriveById(id As Integer) As Boolean
Dim drives As New Collection

Set drives = GetListOfDrives

If Helpers.Exists(drives, id + 1) Then
    CurDrive = id + 1
    
    Ready = True
    SetDriveById = True
Else
    SetDriveById = False
End If
End Function
'*
'* Get a list of all CD-Rom drives in the current computer
'* @return Collection: Collection containing a list of `mdlAdioCdDrive` models with all the drives
'*
Public Function GetListOfDrives() As Collection
Dim a As Long, n As Long
Dim cdi As BASS_CD_INFO
Dim MaxDrives As Integer
Dim ReturnCollection As New Collection
Dim drive As mdlAdioCdDrive

MaxDrives = 10

a = 0
While (a < MaxDrives And BASS_CD_GetInfo(a, cdi) <> 0)
    Set drive = New mdlAdioCdDrive
    
    drive.cdLetter = Chr$(65 + cdi.letter)
    drive.cdDescription = VBStrFromAnsiPtr(cdi.vendor) & " " & VBStrFromAnsiPtr(cdi.product) & " " & VBStrFromAnsiPtr(cdi.rev)

    ReturnCollection.Add drive
    
    RaiseEvent DeviceFound(a + 1, drive.cdDescription, drive.cdLetter)
    
    a = a + 1
Wend

Set GetListOfDrives = ReturnCollection
End Function
'*
'* Open the door of the CD-Rom drive
'*
Public Sub OpenDoor()
Call BASS_CD_Door(CurDrive, BASS_CD_DOOR_OPEN)
RaiseEvent DoorChanged
End Sub
'*
'* Close the door of the CD-Rom drive
'*
Public Sub CloseDoor()
Call BASS_CD_Door(CurDrive, BASS_CD_DOOR_CLOSE)
RaiseEvent DoorChanged
End Sub
'*
'* Resize the usercontrol
'*
Private Sub UserControl_Resize()
width = Image_Main.width
height = Image_Main.height
End Sub
