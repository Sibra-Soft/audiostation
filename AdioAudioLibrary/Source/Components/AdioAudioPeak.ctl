VERSION 5.00
Begin VB.UserControl AdioAudioPeak 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer_Main 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   1215
   End
   Begin VB.Image Image_Main 
      Height          =   480
      Left            =   0
      Picture         =   "AdioAudioPeak.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "AdioAudioPeak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////////
'// FileName        : AdioAudio.ctl
'// FileType        : Microsoft Visual Basic 6 - Usercontrol
'// Author          : Alex van den Berg
'// Created         : 16-08-2023
'// Last Modified   : 30-01-2026
'// Copyright       : Sibra-Soft
'// Description     : Usercontrol for audio peak functionality
'////////////////////////////////////////////////////////////////

Option Explicit

'// Private vars
Private MasterVolume As New clsAdioMasterVolume
Private CurChannel As Long

'// Public vars
Public Bands As Integer
Public Mutted As Boolean

'// Events
Public Event SpectrumLevelChange(col As Integer, Value As Integer)
Public Event ChannelAudioLevelChange(leftValue As Integer, rightValue As Integer)
Public Event MasterAudioLevelChange(Value As Integer)
Public Event MasterAudioPeakLevelChange(Value As Integer)
'*
'* Set the current channel to read from
'* @Param Long TargetChannel The channel you want to read the audio from
'*
Public Sub SetChannel(TargetChannel As Long)
CurChannel = TargetChannel
End Sub
'*
'* Mute the computers master volume
'*
Public Sub MuteMasterVolume()
If Mutted Then
    Call MasterVolume.SetMute(0)
    Mutted = False
Else
    Call MasterVolume.SetMute(1)
    Mutted = True
End If
End Sub
'*
'* Reset spectrum, set all values to 0
'*
Public Sub ResetSpectrum()
Dim X As Integer

For X = 0 To Bands - 1
    RaiseEvent SpectrumLevelChange(X, 0)
Next X
End Sub
'*
'* Start running of the component
'*
Public Sub Run()
Bands = 29

Mutted = MasterVolume.GetMute
RaiseEvent MasterAudioLevelChange(BASS_GetVolume() * 100)

Timer_Main.Enabled = True
End Sub
'*
'* Set the master volume of the computer
'* @param Integer Value: The value to use for the volume (0 - 100)
'*
Public Sub SetMasterVolume(Value As Integer)
Call BASS_SetVolume(Value / 100)
End Sub
'*
'* Timer for keeping track of the audio changes
'*
Private Sub Timer_Main_Timer()
Dim level As Long
Dim Left As Integer
Dim Right As Integer

Dim B0 As Long
Dim Sc As Long, B1 As Long
Dim Sum As Single
Dim X As Integer, Y As Long, y1 As Long
Dim fft(1024) As Single

Dim Endpoint As New clsAdioCoreAudioEndpoint

' Get master audio: Windows Vista and above
RaiseEvent MasterAudioPeakLevelChange(Math.Round(Endpoint.GetPeak * 100))

' Get right and left level
level = BASS_ChannelGetLevel(CurChannel)

Left = (LoWord(level) / 32768) * 100
Right = (HiWord(level) / 32768) * 100

If Left > 100 Then: Exit Sub
If Right > 100 Then: Exit Sub

RaiseEvent ChannelAudioLevelChange(Left, Right)

' Get audio spectrum
Call BASS_ChannelGetData(CurChannel, fft(0), BASS_DATA_FFT2048)

B0 = 0

For X = 0 To Bands - 1
    Sum = 0
    B1 = 2 ^ (X * 10# / (Bands - 1))
    If (B1 > 1023) Then B1 = 1023
    If (B1 <= B0) Then B1 = B0 + 1 ' make sure it uses at least 1 FFT bin
    Sc = 10 + B1 - B0
    Do
        Sum = Sum + fft(1 + B0)
        B0 = B0 + 1

        RaiseEvent SpectrumLevelChange(X, Sum * 100)
    Loop While B0 < B1
Next X
End Sub
Public Sub SetMidiVolume()

End Sub
'*
'* Resize of the usercontrol
'*
Private Sub UserControl_Resize()
Width = Image_Main.Width
Height = Image_Main.Height
End Sub
