VERSION 5.00
Begin VB.UserControl AdioCore 
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   ControlContainer=   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   810
   ScaleWidth      =   1155
   Begin VB.Image Image_Main 
      Height          =   480
      Left            =   0
      Picture         =   "AdioCore.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "AdioCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'///////////////////////////////////////////////////////////////
'// FileName        : AdioCore.ctl
'// FileType        : Microsoft Visual Basic 6 - Usercontrol
'// Author          : Alex van den Berg
'// Created         : 17-08-2023
'// Last Modified   : 28-09-2026
'// Copyright       : Sibra-Soft
'// Description     : Adio core component
'////////////////////////////////////////////////////////////////

'// Public vars
Public AdioChannel As Long

'// Enums
Public Enum enumAdioSupportedFileTypes
    [MID - Midi File] = 0
    [KAR - Karaoke File] = 1
    [SID - Commodore64 8Bit File] = 2
    [MUS - Beep Symfony File] = 3
    [MP2 - MPEG-1 Audio Layer 2 File] = 4
    [MP3 - MPEG-1 Audio Layer 3 File] = 5
    [WAV - Waveform Audio File] = 6
    [WMA - Windows Media Audio] = 7
    [OOG - Digital Multimedia Container File] = 8
    [AIFF - Audio Interchange File] = 9
    [AAC - Advanced Audio Coding File] = 10
    [FLAC - Free Lossless Audio Codec File] = 11
    [CDA - CD Audio] = 12
End Enum

Public Enum enumAdioSeekDirection
    AdioForward
    AdioRewind
End Enum

Public Enum enumAdioFadeType
    AdioIn
    AdioOut
    AdioCross
End Enum

Public Enum enumAdioPlayState
    AdioStopped
    AdioPlaying
    AdioPaused
    AdioEnded
    AdioReady
End Enum

Public Enum enumAdioRepeatMode
    AdioPlayTrack
    AdioRepeatTrack
End Enum

'// Events
Public Event GetVolume(Value As Integer)
Public Event VolumeChanged(Value As Integer)
Public Event Error(Description As String, Code As Integer)
Public Event DeviceFound(id As Integer, Name As String, InputDev As Boolean, OutputDev As Boolean)
Public Event SoundFont(file As String)

'// Private vars
Private Declare Function GetVersion Lib "kernel32" () As Long
Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Form_About.Show vbModal
End Sub
Private Function GetDeviceType(flags As Long) As String
Select Case (flags And BASS_DEVICE_TYPE_MASK)
    Case BASS_DEVICE_TYPE_NETWORK
        GetDeviceType = "Remote Network"
    Case BASS_DEVICE_TYPE_SPEAKERS
        GetDeviceType = "Speakers"
    Case BASS_DEVICE_TYPE_LINE:
        GetDeviceType = "Line"
    Case BASS_DEVICE_TYPE_HEADPHONES:
        GetDeviceType = "Headphones"
    Case BASS_DEVICE_TYPE_MICROPHONE:
        GetDeviceType = "Microphone"
    Case BASS_DEVICE_TYPE_HEADSET:
        GetDeviceType = "Headset"
    Case BASS_DEVICE_TYPE_HANDSET:
        GetDeviceType = "Handset"
    Case BASS_DEVICE_TYPE_DIGITAL:
        GetDeviceType = "Digital"
    Case BASS_DEVICE_TYPE_SPDIF:
        GetDeviceType = "SPDIF"
    Case BASS_DEVICE_TYPE_HDMI:
        GetDeviceType = "HDMI"
    Case BASS_DEVICE_TYPE_DISPLAYPORT:
        GetDeviceType = "DisplayPort"
        
    Case Else
        GetDeviceType = "Unknown"
End Select
End Function
Public Function GetListOfDevices() As Collection
Dim returnList As New Collection
Dim device As mdlAdioDevice
Dim di As BASS_DEVICEINFO
Dim a As Integer

' Get output devices
a = 1
Do While BASS_GetDeviceInfo(a, di)
    Set device = New mdlAdioDevice
    
    device.dId = a
    device.dOutput = True
    device.dDriver = VBStrFromAnsiPtr(di.driver)
    device.dName = VBStrFromAnsiPtr(di.Name)
    device.dtype = GetDeviceType(di.flags)
    
    If (di.flags And BASS_DEVICE_LOOPBACK) Then device.dIsLoopback = True
    If (di.flags And BASS_DEVICE_ENABLED) Then device.dIsEnabled = True
    If (di.flags And BASS_DEVICE_DEFAULT) Then device.dIsDefault = True
    
    a = a + 1
    
    RaiseEvent DeviceFound(device.dId, device.dName, False, True)
    
    returnList.Add device
Loop

' Get input devices
a = 0
Do While BASS_RecordGetDeviceInfo(a, di)
    Set device = New mdlAdioDevice
    
    device.dId = a
    device.dInput = True
    device.dDriver = VBStrFromAnsiPtr(di.driver)
    device.dName = VBStrFromAnsiPtr(di.Name)
    device.dtype = GetDeviceType(di.flags)
    
    If (di.flags And BASS_DEVICE_LOOPBACK) Then device.dIsLoopback = True
    If (di.flags And BASS_DEVICE_ENABLED) Then device.dIsEnabled = True
    If (di.flags And BASS_DEVICE_DEFAULT) Then device.dIsDefault = True
    
    ' Get input devices before Windows Vista
    If (GetVersion() And &HFF) < 6 Then
        ' list inputs
        Dim b As Long
        Dim n As Long
        Call BASS_RecordInit(a)
        b = 0
        Do
            n = BASS_RecordGetInputName(b)
            If n = 0 Then Exit Do

            ' Start adding devices

            b = b + 1
        Loop
        Call BASS_RecordFree
    End If
    
    RaiseEvent DeviceFound(device.dId, device.dName, True, False)
    
    a = a + 1
    
    returnList.Add device
Loop

Set GetListOfDevices = returnList
End Function
Public Sub Initialize()

' Load external library
If (HiWord(BASS_GetVersion) <> BASSVERSION) Then '2.4.10
    Call MsgBox("An incorrect version of BASS.DLL was loaded", vbCritical)
    Exit Sub
End If

' Init sound device
If (BASS_Init(-1, 44100, 0, UserControl.hwnd, 0) = 0) Then
    Call MsgBox("Can't initialize device")
    Exit Sub
End If

' Load plugins
Call BASS_PluginLoad(App.path & "\basscd.dll", 0)
Call BASS_PluginLoad(App.path & "\bassflac.dll", 0)
Call BASS_PluginLoad(App.path & "\basswma.dll", 0)
Call BASS_PluginLoad(App.path & "\basswasapi.dll", 0)

Call GetListOfDevices
End Sub

Private Sub UserControl_Resize()
UserControl.height = Image_Main.height
End Sub

Private Sub UserControl_Terminate()
Call BASS_ChannelStop(AdioChannel)
Call BASS_StreamFree(AdioChannel)
Call BASS_Free
End Sub
