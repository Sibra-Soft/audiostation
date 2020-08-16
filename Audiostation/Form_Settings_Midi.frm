VERSION 5.00
Object = "{291F0443-8437-11CF-840F-444553540000}#1.1#0"; "midifl32.ocx"
Object = "{852E65AD-72F8-11CF-840E-444553540000}#1.1#0"; "midiio32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BF3128D8-55B8-11D4-8ED4-00E07D815373}#1.0#0"; "MBPrgBar.ocx"
Begin VB.Form Form_Settings_Midi 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Midi Settings"
   ClientHeight    =   3735
   ClientLeft      =   3915
   ClientTop       =   4845
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Tag             =   "1022"
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   5280
      Top             =   840
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   7095
      TabIndex        =   5
      Top             =   0
      Width           =   7095
      Begin MBProgressBar.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackPicture     =   "Form_Settings_Midi.frx":0000
         BarPicture      =   "Form_Settings_Midi.frx":001C
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5760
         Top             =   840
      End
      Begin VB.Timer Trm_Main 
         Interval        =   1
         Left            =   4800
         Top             =   840
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Allow playing Midi files (*.mid, *.kar)"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Allow playing Beep Sympony files (*.mus)"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   3975
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Allow playing Commodore64 sound files (*.sid)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4455
      End
      Begin Audiostation.ButtonBig cmdSave 
         Height          =   390
         Left            =   2040
         TabIndex        =   10
         Top             =   3240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   688
         Caption         =   "Save"
      End
      Begin VB.ComboBox OutputDevCombo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   3975
      End
      Begin MidifileLib.Midifile Midifile1 
         Left            =   720
         Top             =   3120
         _Version        =   65537
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   0
         Filename        =   ""
      End
      Begin MidiioLib.MIDIOutput MIDIOutput1 
         Left            =   1320
         Top             =   3120
         _Version        =   65537
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   0
         DeviceID        =   -1
         VolumeLeft      =   65535
         VolumeRight     =   65535
      End
      Begin MidiioLib.MIDIInput MIDIInput1 
         Left            =   120
         Top             =   3120
         _Version        =   65537
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   0
         MessageEventEnable=   -1  'True
      End
      Begin Audiostation.ButtonBig cmdClose 
         Height          =   390
         Left            =   3240
         TabIndex        =   14
         Top             =   3240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   688
         Caption         =   "Close"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mabry Software, Inc."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   1845
         Width           =   2025
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Midi Powered By: "
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1845
         Width           =   1545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Midi output device:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1620
      End
   End
   Begin VB.Timer Timer8 
      Interval        =   100
      Left            =   4680
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   4800
      Top             =   2040
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "\"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2280
         TabIndex        =   3
         Top             =   1440
         Width           =   4185
      End
      Begin VB.Label Label4 
         Caption         =   "0 Sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2280
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "mid"
      DialogTitle     =   "Open MIDI File"
      Filter          =   "(*.mid) MIDI files|*.mid|"
      FilterIndex     =   248
      FontSize        =   2,28347e-38
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "Form_Settings_Midi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VSliderVuDecayValue As Integer
Dim NumVULoaded As Integer

Dim lVolume As Long
Dim rVolume As Long
Dim PreviousTime As Long
Dim TrackOffset As Long

Dim msPerTick As Single
Dim ticksPerMs As Single

Dim CurrentTime As Long

Public Sub CloseOutputDevice()
If MIDIOutput1.State >= MIDISTATE_OPEN Then
    If (MIDIOutput1.HasLRVolume) Then
        MIDIOutput1.VolumeLeft = lVolume
        MIDIOutput1.VolumeRight = rVolume
    ElseIf (MIDIOutput1.HasVolume) Then
        MIDIOutput1.VolumeLeft = lVolume
    End If
    MIDIOutput1.Action = MIDIOUT_CLOSE
End If
End Sub
  
Private Sub DisplayTrackNames()
    Dim m As Integer, maxt As Integer
    Dim T As Integer
    Dim I

    If Midifile1.NumberOfTracks = 1 Then
        TrackOffset = 1
    Else
        TrackOffset = 2
    End If
    maxt = Midifile1.NumberOfTracks
    If maxt > 16 Then maxt = 16
    If NumVULoaded < maxt Then
        For I = NumVULoaded To maxt - TrackOffset


            Next I

        NumVULoaded = maxt - (TrackOffset - 1)
    ElseIf NumVULoaded > maxt Then
        For I = maxt - (TrackOffset - 1) To NumVULoaded - 1

            Next I

        NumVULoaded = maxt - (TrackOffset - 1)
        End If

    DoEvents

    For T = 1 To maxt
        If (T = 1) Then
            msPerTick = ((Midifile1.Tempo) / 1000) / Midifile1.TicksPerQuarterNote
            
            If (Midifile1.Tempo = 0) Then
            
            Else
                ticksPerMs = (Midifile1.TicksPerQuarterNote / Midifile1.Tempo) * 1000
            End If
        End If

        If (T >= 2) Or (Midifile1.NumberOfTracks = 1) Then
            End If
        Next T
End Sub

Private Sub cmdClose_Click()
Hide
End Sub

Private Sub cmdSave_Click()
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "MidiOutputDevice", OutputDevCombo.ListIndex)

Hide
End Sub

Private Sub Form_Load()
Dim I As Integer

Midifile1.Tempo = 100
NumVULoaded = 1
VSliderVuDecayValue = 15

If MIDIOutput1.DeviceCount = 0 Then

Else
    For I = -1 To MIDIOutput1.DeviceCount - 1
        MIDIOutput1.DeviceID = I
        OutputDevCombo.AddItem MIDIOutput1.ProductName
    Next
    
    OutputDevCombo.ListIndex = Settings.ReadSetting("Sibra-Soft", "Audiostation", "MidiOutputDevice", 0)
End If

Check1.Value = Settings.ReadSetting("Sibra-Soft", "Audiostation", "PluginBeepSymphony", 1)
Check2.Value = Settings.ReadSetting("Sibra-Soft", "Audiostation", "PluginBeepBox", 1)
Check4.Value = Settings.ReadSetting("Sibra-Soft", "Audiostation", "PluginMidi", 1)
End Sub
   
Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Hide
End Sub

Private Function GetTrackName(Track As Integer) As String
    Dim I As Integer, bnk As Integer, map As Integer

    Midifile1.TrackNumber = Track
    bnk = 0: map = 0
    For I = 1 To Midifile1.MessageCount
        Midifile1.MessageNumber = I
        '
        'Meta Event
        '
        If (Midifile1.Message = 255) And (Midifile1.Data1 = 3) Then
            If (Midifile1.MsgText = "") Then
                GetTrackName = "Track" & str(Track) & " (null)"
            Else
                GetTrackName = Midifile1.MsgText
            End If
        End If
        If (Midifile1.Message >= &HB0 And Midifile1.Data1 = &H0) Then
            bnk = Midifile1.Data2
        End If
        If (Midifile1.Message >= &HB0 And Midifile1.Data1 = &H20) Then
            map = Midifile1.Data2
        End If
        If (Midifile1.Message >= &HC0 And Midifile1.Message < &HD0) Then
            ' Use next line if desired :)
            'GetTrackName = "Channel " + Str$(Midifile1.Message - &HC0 + 1) + " - Patch:" + Str$(1 + Midifile1.Data1) + " Bank/Map:" + Str$(bnk) + "/" + Trim$(Str$(map))
            Exit Function
        End If
        Next I
    ' GetTrackName = "Channel " + Str$(1 + Midifile1.Message And &HF) + " - No Patch"
End Function

Private Sub Midifile1_Error(ErrorCode As Integer, ErrorMessage As String)
MsgBox ErrorMessage, vbCritical, ErrorCode
End Sub

Private Sub MIDIOutput1_MessageSent(MessageTag As Long)
Dim TrackNumber As Integer
Dim Intensity As Integer
   
If (MessageTag < 0) Or (MessageTag >= 16000) Then
    Exit Sub
End If

Intensity = MessageTag Mod 1000
TrackNumber = Int(MessageTag / 1000)

If (Intensity > 0) And (TrackNumber > 0) And (Intensity > Form_Main.VU_Midi(TrackNumber - TrackOffset).Position) Then
    Form_Main.VU_Midi(TrackNumber - TrackOffset).Position = Intensity
End If
End Sub

Private Sub MIDIOutput1_QueueEmpty()
StopPlay
End Sub
Private Sub MidiReset()
    Dim X As Integer
    Dim Y As Integer

    For Y = 0 To NumVULoaded - 1
        
        Next Y

    For Y = 0 To 15
        For X = 0 To 127
            ' Turn all note off
            MIDIOutput1.Data1 = X
            MIDIOutput1.Data2 = 0
            MIDIOutput1.Action = MIDIOUT_SEND
            Next
        Next
End Sub

Private Sub OpenOutputDevice()
MIDIOutput1.DeviceID = OutputDevCombo.ListIndex - 1
MIDIOutput1.Action = MIDIOUT_OPEN
    
If (MIDIOutput1.HMidiDevice <> 0) Then
If (MIDIOutput1.HasLRVolume) Then
    lVolume = MIDIOutput1.VolumeLeft
    rVolume = MIDIOutput1.VolumeRight
Else
    If (MIDIOutput1.HasVolume) Then
        lVolume = MIDIOutput1.VolumeLeft
    End If
End If
End If
End Sub
Private Sub OutputDevCombo_Click()
StopPlay
End Sub
Private Sub QueueSong()
    Dim m As Integer
    Dim mm As Integer
    Dim Track As Integer
    Dim I As Integer
    ReDim CurrentTimeQueue(Midifile1.NumberOfTracks) As Long
    ReDim PreviousTimeQueue(Midifile1.NumberOfTracks) As Long
    ReDim LowestEvent(Midifile1.NumberOfTracks) As Long
    ReDim TrackDone(Midifile1.NumberOfTracks) As Integer
    Dim TracksLoadComplete As Integer
    Dim IncrementAmount As Integer
    Dim OldMousePointer As Integer
    Dim MessageCount As Long
    Dim MessageTotal As Long

    If (Midifile1.FileName = "") Then
        Exit Sub
    End If
    DoEvents
    MIDIOutput1.Action = MIDIOUT_RESET

    OldMousePointer = Screen.MousePointer
    Screen.MousePointer = 11

    MessageTotal = 0
    For m = 1 To Midifile1.NumberOfTracks
        LowestEvent(m) = 1
        TrackDone(m) = False
        Midifile1.TrackNumber = m
        MessageTotal = MessageTotal + Midifile1.MessageCount
    Next m
    
    If Midifile1.NumberOfTracks = 1 Then
        TracksLoadComplete = 0
        TrackOffset = 1
    Else
        TracksLoadComplete = 1
        TrackOffset = 2
        End If

    IncrementAmount = 125
    MessageCount = 0

    Do While TracksLoadComplete < Midifile1.NumberOfTracks
        If MessageTotal <> 0 Then

        End If
            
        For Track = TrackOffset To Midifile1.NumberOfTracks
            Midifile1.TrackNumber = Track
            
            If TrackDone(Track) = False Then
                'Increment throught in groups of IncrementAmount events
                If Midifile1.MessageCount > LowestEvent(Track) + IncrementAmount Then
                    mm = LowestEvent(Track) + IncrementAmount
                Else
                    mm = Midifile1.MessageCount
                End If
    
                For m = LowestEvent(Track) To mm
                    Midifile1.MessageNumber = m
                
                    ' Put message data in control
                    MIDIOutput1.Message = Midifile1.Message
                    MIDIOutput1.Data1 = Midifile1.Data1
                    MIDIOutput1.Data2 = Midifile1.Data2
    
                    'Tag notes to play on keyboard and VU meters
                    If (Midifile1.Message And &HF0) = note_aan Then
                        MIDIOutput1.MessageTag = Midifile1.Data2 + 1 + (Track * 1000)
                    Else
                        MIDIOutput1.MessageTag = 0
                        End If
    
                    CurrentTimeQueue(Track) = PreviousTimeQueue(Track) + Midifile1.time
    
                    MIDIOutput1.time = Int(CurrentTimeQueue(Track) * msPerTick)
                    PreviousTimeQueue(Track) = CurrentTimeQueue(Track)
                
                    ' Add to output queue
                    MIDIOutput1.Action = MIDIOUT_QUEUE
                    MessageCount = MessageCount + 1
                Next m

                If mm = Midifile1.MessageCount Then
                    TrackDone(Track) = True
                    TracksLoadComplete = TracksLoadComplete + 1
                Else
                    LowestEvent(Track) = LowestEvent(Track) + IncrementAmount + 1
                    End If
                End If
            Next
        Loop
        
ProgressBar1.Value = 0

If (MIDIOutput1.time / 1000 > ProgressBar1.Max) Then
    ProgressBar1.Max = MIDIOutput1.time / 1000
End If

Timer4.Enabled = True

Screen.MousePointer = OldMousePointer
End Sub

Public Sub StartPlay(FileName As String)
Dim K As Integer

For K = 0 To Form_Main.VU_Midi.Count - 1
    Form_Main.VU_Midi(K).Position = 0
Next

CloseOutputDevice
StopPlay

Midifile1.Action = MIDIFILE_CLEAR
Midifile1.FileName = FileName
Midifile1.Action = MIDIFILE_OPEN

DisplayTrackNames
DoEvents

OpenOutputDevice
QueueSong
MIDIOutput1.Action = MIDIOUT_START
Timer1.Enabled = True

'Set Midi Control Set
With Form_Main
    .Command16.Enabled = True           'Previous Track
    .Command15.Enabled = True           'Rewind Track
    .Command14.Enabled = True           'Stop
    .Command13.Enabled = True           'Play
    .Command10.Enabled = True           'Pause
    .Command12.Enabled = True           'Forward Track
    .Command9.Enabled = True            'Next Track
End With
End Sub
Public Sub StopPlay()
MIDIOutput1.Action = MIDIOUT_STOP
MidiReset
CloseOutputDevice
End Sub
Private Sub Timer1_Timer()
Dim n As Integer

For n = 0 To NumVULoaded - 1
    If Form_Main.VU_Midi(n).Position > VSliderVuDecayValue Then
        Form_Main.VU_Midi(n).Position = Form_Main.VU_Midi(n).Position - VSliderVuDecayValue
    Else
        Form_Main.VU_Midi(n).Position = 0
    End If
Next n
End Sub
Private Sub Timer4_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1

If ProgressBar1.Percentage = 100 Then
    Me.StopPlay
    
    Form_Main.Trm_Lights_Midi.Tag = 2
    Form_Main.Trm_Midi_Play.Enabled = False
End If
End Sub

Private Sub Timer8_Timer()
DisplayTrackNames
DoEvents
Timer8.Enabled = False
End Sub

Private Sub Trm_Main_Timer()
If Check4.Value = vbChecked Then
    OutputDevCombo.Enabled = True
Else
    OutputDevCombo.Enabled = False
End If
End Sub
