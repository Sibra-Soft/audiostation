VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0518EEBD-7F0E-4513-8491-A0221C9008A2}#2.1#0"; "midiio2k.ocx"
Object = "{4424C993-EABF-4A03-9BA9-369E0F07466E}#1.2#0"; "midifl2k.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form_Midi 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Midi Player"
   ClientHeight    =   3615
   ClientLeft      =   300
   ClientTop       =   1110
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
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
   ScaleHeight     =   3615
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.TextBox TextIndicator1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "TextIndicator1"
      Top             =   2340
      Width           =   3135
   End
   Begin VB.CheckBox CheckAutoStop 
      Caption         =   "Auto Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   28
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer TimerEnd 
      Enabled         =   0   'False
      Left            =   420
      Top             =   2460
   End
   Begin VB.CommandButton CommandSetFocus 
      Caption         =   "CommandSetFocus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer TimerMidiError 
      Enabled         =   0   'False
      Left            =   300
      Top             =   2400
   End
   Begin ComctlLib.ProgressBar VIndicatorPeak 
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   5
      Top             =   2580
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   128
   End
   Begin VB.CommandButton CommandIndicatorPeak 
      BackColor       =   &H80000002&
      Caption         =   "CommandIndicatorPeak"
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2580
      Width           =   795
   End
   Begin ComctlLib.ProgressBar VIndicator1 
      Height          =   150
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   265
      _Version        =   327682
      Appearance      =   1
      Max             =   128
   End
   Begin MidifileLib.MIDIFile MIDIFile1 
      Left            =   1680
      Top             =   2340
      _Version        =   65538
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Filename        =   ""
   End
   Begin MidiioLib.MIDIOutput MIDIOutput1 
      Left            =   1140
      Top             =   2340
      _Version        =   131073
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Timer TimerProgressBar 
      Enabled         =   0   'False
      Left            =   180
      Top             =   2340
   End
   Begin VB.Frame Frame2 
      Caption         =   "Playback Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   6060
         TabIndex        =   7
         Top             =   120
         Width           =   2055
         Begin VB.CheckBox CheckAutoReplay 
            Caption         =   "Auto Replay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox CheckMidiOutFilterLateEventAllMax 
            Caption         =   "Filter all Notes if late"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1260
            Width           =   1755
         End
         Begin VB.CheckBox CheckMidiOutFilterFF 
            Caption         =   "Filter FF optimized"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   1755
         End
         Begin VB.HScrollBar HScrollPlayerTime 
            Height          =   255
            Left            =   180
            TabIndex        =   16
            Top             =   180
            Width           =   1815
         End
         Begin VB.CommandButton CmdStop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Stop"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   19
            Top             =   660
            Width           =   615
         End
         Begin VB.CommandButton CmdPause 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Pause"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   18
            Top             =   660
            Width           =   615
         End
         Begin VB.CommandButton CmdPlay 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Play"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   660
            Width           =   615
         End
         Begin VB.Label LabelQueueTime 
            Alignment       =   2  'Center
            Caption         =   "LabelQueueTime"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   420
            Width           =   1815
         End
      End
      Begin VB.CheckBox CheckManualSort 
         Caption         =   "Force sort after open"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3960
         TabIndex        =   15
         Top             =   1500
         Width           =   1995
      End
      Begin VB.OptionButton OptionOpen 
         Caption         =   "Open as midi 0 in tracks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3960
         TabIndex        =   14
         Top             =   1320
         Width           =   1995
      End
      Begin VB.OptionButton OptionOpen 
         Caption         =   "Open as midi 0 in pieces"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3960
         TabIndex        =   13
         Top             =   1140
         Width           =   2055
      End
      Begin VB.CommandButton CmdOpen 
         Caption         =   "Open File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton OptionOpen 
         Caption         =   "Open as midi 1 in tracks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3960
         TabIndex        =   12
         Top             =   960
         Width           =   1995
      End
      Begin VB.CheckBox CheckPatch 
         Caption         =   "Always Display Patch Names"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   900
         Width           =   2595
      End
      Begin VB.ComboBox OutputDevCombo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   540
         Width           =   2595
      End
      Begin VB.CheckBox CheckPauseRestart 
         Caption         =   "Pause can restart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1740
         TabIndex        =   31
         Top             =   1620
         Width           =   1695
      End
      Begin VB.CheckBox CheckTextIndicator 
         Caption         =   "Always Display Mnemonic Indicator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   2955
      End
      Begin VB.CheckBox CheckBarDecay 
         Caption         =   "Bar Decay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CheckBox CheckPeakHold 
         Caption         =   "Peak Hold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   1155
      End
      Begin VB.CheckBox CheckPeakDecay 
         Caption         =   "Peak Decay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Tracks:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4020
         TabIndex        =   27
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label LabelNumberOfTracks 
         Caption         =   "LabelNumberOfTracks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   26
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "File name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   795
      End
      Begin VB.Label LabelPlayerFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LabelPlayerFile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Output Device:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2220
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "mid"
      DialogTitle     =   "Open MIDI File"
      Filter          =   "(*.mid) MIDI files|*.mid|"
      FilterIndex     =   248
      FontSize        =   2,28347e-38
   End
   Begin VB.Label LabelTrackName 
      Caption         =   "Track Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form_Midi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VSliderVuBarDecay As Single
Dim VSliderVuPeakHold As Single
Dim VSliderVuPeakDecay As Single
Dim NumVULoaded As Integer

Dim PreviousTime As Long
Dim TrackOffset As Long
Dim MainStreamNumber As Integer
Dim MainStreamGroup() As Integer
Dim MainStreamOption As Integer
Dim MainMidifile As String
Dim TimeExpectedMessage As Long
Dim TimeExpectedMessageRelToTempo As Long
Dim TimeExpectedMessageRelToOpen As Long
Dim TimeActualMessageRelToOpen As Long

Dim msPerTick As Single
Dim ticksPerMs As Single
Dim CurrentTime As Long
Dim maxt As Integer
Dim TrackVis(255) As Integer

Const MB_OPTIONOPENDEFAULT = 0
Const MB_STREAMNUMBER = 1
Const MB_STREAMEMPTY = 2

Const MB_STREAMNAME_1 = "stream" ' same for all in midi 1
Const MB_STREAMNAME_FF = "FFstream"

Const MB_HSCROLLTIMESCALEOFFSET = 1000& ' max scroll = 32767000msec, min scroll = 1000msec
Const MB_HSCROLLMESSAGESCALEOFFSET = 10& ' max scroll = 327670messages, min scroll = 10messages

Private Sub CheckAutoReplay_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close
    
    Dim mGroupNumber As Integer
    
    ' Need to change in real time too, before/after, queue/start.

    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative

    If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
        ' Midi format 0
        If MainStreamNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamNumber
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            Else
                MIDIOutput1.StreamAutoReplay = IIf(CheckAutoReplay.value = 0, False, True)
            End If
        End If
    
    Else
        ' Midi format 1
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            Else
                MIDIOutput1.StreamAutoReplay = IIf(CheckAutoReplay.value = 0, False, True)
            End If
        Next mGroupNumber
    End If
    
ExitSection:
    
    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub CheckAutoStop_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close
    
    Dim mGroupNumber As Integer
    
    ' Need to change in real time too, before/after, queue/start.
    
    ' Always stop by default, in case no other options are set
    If CheckAutoStop.value <> 1 Then _
     CheckAutoStop.value = 1
    
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    
    If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
        ' Midi format 0
        If MainStreamNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamNumber
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            Else
                MIDIOutput1.StreamAutoStop = IIf(CheckAutoStop.value = 0, False, True)
            End If
        End If
    
    Else
        ' Midi format 1
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            Else
                MIDIOutput1.StreamAutoStop = IIf(CheckAutoStop.value = 0, False, True)
            End If
        Next mGroupNumber
    End If

ExitSection:

    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub CheckBarDecay_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Need to change in real time too, before/after, queue/start.

    ' On/Off controlled with decay value.
    ' Set always some decay value or else
    ' meters may flicker changes from a Note-On
    ' to Note-Off much too fast while playing.
    ' (range 1 to 25, increment in Vindicator1.Max and TimerProgressBar.Interval)
    Select Case CheckBarDecay.value
     Case 0: VSliderVuBarDecay = 25 ' no decay since very fast
     Case 1: VSliderVuBarDecay = 10 ' fast decay
    End Select
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub CheckMidiOutFilterLateEventAllMax_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close
    
    ' Need to change in real time too, before/after, queue/start.
    
    MIDIOutput1.FilterLateEventAllMax = IIf(CheckMidiOutFilterLateEventAllMax.value = 0, False, True)

ExitSection:

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub CheckPatch_Click()
    '
End Sub

Private Sub CheckPauseRestart_Click()
    '
End Sub

Private Sub CheckPeakHold_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim i As Integer
    
    ' Need to change in real time too, before/after, queue/start.

    ' WARNING,
    ' The peakhold value may decay sooner than expected
    ' since calculations later do not always track the
    ' latest values accurately.

    ' On/Off controlled with checkbox and visible later.
    ' (range 1 to X msec)
    VSliderVuPeakHold = 1500
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub CheckPeakDecay_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Need to change in real time too, before/after, queue/start.

    ' On/Off controlled with checkbox and visible later.
    ' (range 1 to 25, increment in Vindicator1.Max and TimerProgressBar.Interval)
    VSliderVuPeakDecay = 10 ' fast decay
    'VSliderVuPeakDecay = 5 ' slow decay
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub CmdPause_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    PausePlay
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub CmdPlay_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    StartPlay
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub CmdStop_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    StopPlay
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub CommandSetFocus_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Prevent moving the focus to an object in the default
    ' taborder of the form for safety. Safe to continuously
    ' press enter on a control without inadvertantly
    ' cascading through other buttons on the form.

    CommandSetFocus.top = Me.top + Me.Height - 1000 ' hidden from form but still visible
    CommandSetFocus.Width = 0 ' smaller in case not hidden properly
    CommandSetFocus.TabStop = False ' hidden from access

    If CommandSetFocus.Enabled = False Then CommandSetFocus.Enabled = True ' must always be available
    If CommandSetFocus.Visible = False Then CommandSetFocus.Visible = True ' must always be available
    
    If Me.WindowState <> vbMinimized And Me.Visible = True Then
        ' Move focus to hiden object
        ' and stay on that object.
        CommandSetFocus.SetFocus
    End If

    ' Alternative,
    ' object.Enable = False ' loose the focus to next object in taborder
    ' object.Enable = True ' reenable so still accessible later
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

 
Public Function OpenFile(Filename As String) As Boolean
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Prepare close
    Call ClearPlayer

    ' Prepare midi file
    MIDIFile1.action = MIDIFILE_CLEAR
    MIDIFile1.action = MIDIFILE_CLOSE
    MIDIFile1.ReadOnly = True ' avoid sharing violations when loading files
    MIDIFile1.Filename = Filename
    MIDIFile1.action = MIDIFILE_OPEN
    DisplayTrackNames
    MainMidifile = MIDIFile1.Filename
    Me.Caption = MainMidifile
    If Len(MIDIFile1.Filename) <= 40 Then
        LabelPlayerFile.Caption = MIDIFile1.Filename
    Else
        LabelPlayerFile.Caption = ". . . " & Right(MIDIFile1.Filename, 40)
    End If
    
    ' Prepare midi data
    MIDIOutput1_Initialize
    If OptionOpen(MB_OPTIONOPENDEFAULT).value = True Then QueueSong_ByMidi1Track
    If OptionOpen(1).value = True Then QueueSong_ByMidi0Pieces
    If OptionOpen(2).value = True Then QueueSong_ByMidi0Tracks
    'If CheckManualSort.Value = ... already implemented
    If OptionOpen(MB_OPTIONOPENDEFAULT).value = True Then MainStreamOption = 0
    If OptionOpen(1).value = True Then MainStreamOption = 1
    If OptionOpen(2).value = True Then MainStreamOption = 2
    Call ClearScrollBar
    Call DisplayPlayerButtons
   
    ' Midi file not needed anymore
    MIDIFile1.action = MIDIFILE_CLEAR
    MIDIFile1.action = MIDIFILE_CLOSE
    MIDIFile1.Filename = ""
    
    Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents

ExitSection:
    gisCurrentQueue = False ' not needed anymore

    Exit Function
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Function
   
Private Sub Form_ZOrder()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Default zorder of objects on screen to align borders (zorder 0)
    ' textbox & checkbox is in reverse order to align borders (zorder 1)
    ' (1 is to back of layer, 0 is to front of layer)
    ' (not sure about frames and tabs)
    
    'Const MB_ZORDER_Default = 0 ' clutters the code, so only show exceptions and more brief for brevity
    Const MB_ZORDER_TextAndCheckbox = 1
    Const MB_ZORDER_Front = 0
    Const MB_ZORDER_Back = 1
    
    'CmdOpen.ZOrder 0
    'etc.
    ' (just set order manually on form at design time)
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub Form_TabIndex()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Default taborder of objects on screen (tabindex = 999)
    ' Reverse order (tabindex = 0)
    
    CmdOpen.TabIndex = 999 ' first default object
    OptionOpen(0).TabIndex = 999
    OptionOpen(1).TabIndex = 999
    OptionOpen(2).TabIndex = 999
    CheckManualSort.TabIndex = 999
    HScrollPlayerTime.TabIndex = 999
    CmdPlay.TabIndex = 999
    CmdPause.TabIndex = 999
    CmdStop.TabIndex = 999
    CheckAutoReplay.TabIndex = 999
    CheckAutoStop.TabIndex = 999
    CheckMidiOutFilterLateEventAllMax.TabIndex = 999
    CheckMidiOutFilterFF.TabIndex = 999
    CheckPatch.TabIndex = 999
    CheckTextIndicator.TabIndex = 999
    CheckBarDecay.TabIndex = 999
    CheckPeakHold.TabIndex = 999
    CheckPeakDecay.TabIndex = 999
    CheckPauseRestart.TabIndex = 999
    OutputDevCombo.TabIndex = 999
    CommandSetFocus.TabIndex = 999: CommandSetFocus.TabStop = False ' hidden from access
    
    'VIndicator1(0)
    'VIndicatorPeak(0)
    CommandIndicatorPeak(0).TabIndex = 999: CommandIndicatorPeak(0).TabStop = False ' hidden from access
    TextIndicator1(0).TabIndex = 999: TextIndicator1(0).TabStop = False ' readonly
    
    'CmdOpen.TabIndex = 0 ' not applicable to rearrange first default object
    'CmdOpen.SetFocus ' not applicable at Form_Load()

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub Form_Load()
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown

Dim i As Integer

gmThreadPriorityApp = 1

Call SetThreadPriority(GetCurrentThread(), gmThreadPriorityApp)

If GetThreadPriority(GetCurrentThread()) <> 1 Then Debug.Print "Testing thread priority, " & Trim$(str$(GetThreadPriority(GetCurrentThread())))

' Prepare for errors
MIDIOutput1.ErrorScheme = 1 ' GetFirstError()
MIDIOutput1.ErrorHalt = True ' stop on all errors, easier to troubleshoot

If MIDIOutput1.ErrorScheme <> 1 Then Debug.Print "testing errorscheme"
If MIDIOutput1.ErrorHalt <> False Then Debug.Print "testing errorhalt"

' Fill output device combo box
For i = -1 To MIDIOutput1.DeviceCount - 1
    MIDIOutput1.DeviceID = i
    OutputDevCombo.AddItem MIDIOutput1.ProductName
Next i

' Select first device in list
OutputDevCombo.ListIndex = Settings.ReadSetting("Sibra-Soft", "Audiostation", "MidiPlaybackDeviceId", 0)

Me.top = 10
Me.Left = (Screen.Width - Me.Width) / 2
NumVULoaded = 1

Call DisplayOneTrack(0) ' adjust default progress bar

CheckBarDecay.value = 1: Call CheckBarDecay_Click
CheckPeakHold.value = 1: Call CheckPeakHold_Click
CheckPeakDecay.value = 1: Call CheckPeakDecay_Click
CheckPauseRestart.value = 0: Call CheckPauseRestart_Click

OptionOpen(MB_OPTIONOPENDEFAULT).value = True
OptionOpen(1).value = False
OptionOpen(2).value = False
CheckManualSort.value = 0 ': Call CheckManualSort_Click
HScrollPlayerTime.LargeChange = 10 ' ten seconds
HScrollPlayerTime.SmallChange = 1 ' one second (1000 msec)
HScrollPlayerTime.max = 0 ' determine when open
HScrollPlayerTime.min = 0
CheckMidiOutFilterLateEventAllMax.value = 1: Call CheckMidiOutFilterLateEventAllMax_Click
CheckMidiOutFilterFF.value = 1 ': Call CheckMidiOutFilterFF_Click

TimerProgressBar.Interval = 55 ' part of progress bar decay calculations
TimerProgressBar.Enabled = True
TimerMidiError.Interval = 100 ' fast enough to show unexpected errors
TimerMidiError.Enabled = True
'TimerEnd.Interval = ...
TimerEnd.Enabled = False ' only need if loaded again

Call Form_ZOrder
Call Form_TabIndex

Call ClearPlayer

Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
    If gisEnd = True Then
        ' May have reopened and initialized if processes were
        ' running in the background during shutdown.
        ' Rely on timers to shutdown.
        gisEnd = False ' reset since not shutdown yet
        TimerEnd.Interval = 1 ' as fast as possible
        TimerEnd.Enabled = True ' only need if loaded again
        'Unload Me ' not applicable if background processes still running
    End If
End Sub
   
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Force unload in case logging out of or
    ' shutting down windows.
    ' - in case close with X at corner of window of main form (unload() may fail to run)
    ' - in case right-clicking program in windows taskbar (unload() may fail to run)
    ' - in case logging out of windows (unload() may fail to run)
    ' - in case shutting down some other way (unload() may fail to run)
    Cancel = False
    Call Form_Unload(False)
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    MIDIOutput1_Unload
    
    ' Unload, release and unwind multithreading and
    ' background processes in events and timers.
    ' Assume each procedure has ExitEnd to terminate without stopping.
    ' Assume each procedure has OnError to terminate without stopping.
    ' Assume FormLoad runs timer to unload again as soon as possible.
    gisEnd = True
        
    ' Force to exit if not unloading, releasing, and unwinding
    ' multithreading and background process in events, timers,
    ' and other open forms. Otherwise not recommended.
    'End ' halt
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub ClearTrackBars()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim X As Integer
    Dim Y As Integer

    For Y = 0 To NumVULoaded - 1 ' 0-based scale
        'LabelTrackName(y).Caption = ""
        VIndicator1(Y).value = 0
        VIndicatorPeak(Y).value = 0
        'CommandIndicatorPeak(y)
        TextIndicator1(Y).Text = ""
    Next Y
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Function GetTrackName(track As Integer) As String
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim i As Long, bnk As Integer, map As Integer
    Dim s1 As String

    MIDIFile1.TrackNumber = track
    bnk = 0: map = 0: TrackVis(track) = 1
    For i = 1 To MIDIFile1.MessageCount ' 1-based scale
        MIDIFile1.MessageNumber = i ' 1-based scale
        '
        'Meta Event
        '
        If (MIDIFile1.Message = 255) And (MIDIFile1.Data1 = 3 Or MIDIFile1.Data1 = 1) Then
            If (MIDIFile1.MsgText = "") Then
                GetTrackName = "Track" & str(track) & " (null)"
            Else
                If GetTrackName = "" Then
                    GetTrackName = MIDIFile1.MsgText
                End If
            End If
        End If
        If (MIDIFile1.Message >= &HB0 And MIDIFile1.Data1 = &H0) Then
            bnk = MIDIFile1.Data2
        End If
        If (MIDIFile1.Message >= &HB0 And MIDIFile1.Data1 = &H20) Then
            map = MIDIFile1.Data2
        End If
        If (MIDIFile1.Message >= &HC0 And MIDIFile1.Message < &HD0) Then
            ' Use next line if desired :)
            s1 = "Channel " + str$(MIDIFile1.Message - &HC0 + 1) _
            + " - Patch: " + str$(1 + MIDIFile1.Data1) _
            + "    Bank/Map: " + str$(bnk) + "/" + Trim$(str$(map))
            If GetTrackName = "" Or CheckPatch.value = vbChecked Then
                GetTrackName = s1
            End If
            Exit Function
        End If
    Next i
    If GetTrackName = "" And MIDIFile1.Message <> 255 Then
        GetTrackName = "Channel " + str$(1 + MIDIFile1.Message And &HF) + " - No Patch"
    End If
    If MIDIFile1.Message = 255 Then ' empty track
        TrackVis(track) = 0
    End If
    Exit Function
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Function

Private Sub DisplayTrackNames()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim m As Integer
    Dim T As Integer
    Dim i As Integer

    If MIDIFile1.NumberOfTracks = 1 Then
        TrackOffset = 1
    Else
        TrackOffset = 2
    End If
    maxt = MIDIFile1.NumberOfTracks
    If maxt > 32 Then maxt = 32
    
    Call DisplayOneTrack(0) ' adjust default progress bar

    If NumVULoaded < maxt Then
        ' Add more progress bar to match tracks
        For i = NumVULoaded To maxt - TrackOffset
            Call DisplayOneTrack(i)
        Next i

        NumVULoaded = maxt - (TrackOffset - 1)
    
    ElseIf NumVULoaded > maxt Then
        ' Remove more progress bar to match tracks
        For i = maxt - (TrackOffset - 1) To NumVULoaded - 1
            Unload LabelTrackName(i)
            Unload VIndicator1(i)
            Unload VIndicatorPeak(i)
            Unload CommandIndicatorPeak(i)
            Unload TextIndicator1(i)
        Next i

        NumVULoaded = maxt - (TrackOffset - 1)
    End If

    ' Update Screen.
    T = 200
    If maxt > 16 Then T = 200
    Me.Height = (Me.Height - Me.ScaleHeight) + (NumVULoaded * T) + 2100
    Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents

    For T = 1 To maxt
        'If (t = 1) Then
        '    msPerTick = ((MIDIFile1.Tempo) / 1000) / MIDIFile1.TicksPerQuarterNote
        '    ticksPerMs = (MIDIFile1.TicksPerQuarterNote / MIDIFile1.Tempo) * 1000
        'End If

        If (T >= 2) Or (MIDIFile1.NumberOfTracks = 1) Then
            LabelTrackName(T - TrackOffset).Caption = Trim(GetTrackName(T))
            If TrackVis(T) = 0 Then
                If T > 2 Then
                    'LabelTrackName(y).Visible = False ' not applicable
                    VIndicator1(T - 2).Visible = False
                    VIndicatorPeak(T - 2).Visible = False
                    CommandIndicatorPeak(T - 2).Visible = False
                    TextIndicator1(T - 2).Visible = False
                End If
            Else
                If T > 2 Then
                    'LabelTrackName(y).Visible = True ' not applicable
                    VIndicator1(T - 2).Visible = True
                    'VIndicatorPeak(t - 2).Visible = True ' not yet
                    'CommandIndicatorPeak(t - 2).Visible = True ' not yet
                    'TextIndicator1(t - 2).Visible = True ' not yet
                End If
            End If
        End If
    Next T

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub DisplayOneTrack(ByVal mIndex As Integer)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim i As Integer, xx As Integer, yy As Integer
    Dim s1 As String
    
    i = mIndex
    If i = 0 Then ' first time only
        If maxt < 16 Then
            xx = 200
            LabelTrackName(0).Font = "MS San Serif"
            LabelTrackName(0).FontSize = 8
            LabelTrackName(0).Height = 200
            ' s1 = s1
        Else ' small fonts
            xx = 200
            'LabelTrackName(0).Font = "Small Fonts"
            LabelTrackName(0).FontSize = 8
            LabelTrackName(0).Height = 200
        End If
        TextIndicator1(0).top = VIndicator1(0).top
        TextIndicator1(0).Left = VIndicator1(0).Left
    End If
    xx = LabelTrackName(0).Height
    If i <> 0 Then
        ' Adjust new progress bar.
        ' Inherits properties of object at index = 0.
        Load LabelTrackName(i)
        LabelTrackName(i).top = 2000 + i * xx
        LabelTrackName(i).Left = LabelTrackName(0).Left
        LabelTrackName(i).Caption = ""
        LabelTrackName(i).Font = LabelTrackName(0).Font

        Load VIndicator1(i)
        VIndicator1(i).top = 2050 + i * xx
        VIndicator1(i).Left = VIndicator1(0).Left

        Load VIndicatorPeak(i)
        Load CommandIndicatorPeak(i)
        
        Load TextIndicator1(i)
        TextIndicator1(i).top = VIndicator1(i).top
        TextIndicator1(i).Left = VIndicator1(i).Left
        TextIndicator1(i).Text = ""
        TextIndicator1(i).Font = TextIndicator1(i).Font
    
    ElseIf i = 0 Then
        ' Prepare default progress bar
    End If
        
    ' Adjust default and new progress bar
    
    'LabelTrackName(i).Top = ... ' already positioned on form
    
    'Vindicator1(i).Top = ... ' already positioned on form
    VIndicator1(i).value = 0
    VIndicator1(i).max = 128
    
    VIndicatorPeak(i).top = VIndicator1(i).top ' must match other progress bar
    VIndicatorPeak(i).Left = VIndicator1(i).Left ' must match other progress bar
    VIndicatorPeak(i).Height = VIndicator1(i).Height / 2 ' different in case accidently visible
    VIndicatorPeak(i).Width = VIndicator1(i).Width ' must match other progress bar
    VIndicatorPeak(i).value = VIndicator1(i).value ' must match other progress bar
    VIndicatorPeak(i).max = VIndicator1(i).max ' must match other progress bar
    
    CommandIndicatorPeak(i).top = VIndicator1(i).top ' must match other progress bar
    CommandIndicatorPeak(i).Left = VIndicator1(i).Left ' must match other progress bar
    CommandIndicatorPeak(i).Height = VIndicator1(i).Height ' must match other progress bar
    CommandIndicatorPeak(i).Width = 15
    CommandIndicatorPeak(i).Visible = False ' not visible yet
    CommandIndicatorPeak(i).Enabled = False ' never used
    CommandIndicatorPeak(i).BackColor = &H80000002
    CommandIndicatorPeak(i).TabStop = False ' never used
    If CommandIndicatorPeak(i).Style <> 1 Then Err.Raise 1, , "" ' graphical

    LabelTrackName(i).Visible = True
    VIndicator1(i).Visible = True
    VIndicatorPeak(i).Visible = False ' hidden
    CommandIndicatorPeak(i).Visible = False ' not yet
    TextIndicator1(i).Visible = False ' not yet

#If 1 = 0 Then ' comment out to enable test
    ' Verify the peak calculations are synchronized
    ' (may have to reposition it if no room to fit on screen)
    VIndicatorPeak(i).Visible = True
#End If

    ' ZOrder
    LabelTrackName(i).ZOrder 0
    VIndicator1(i).ZOrder 0
    VIndicatorPeak(i).ZOrder 0
    CommandIndicatorPeak(i).ZOrder 0
    TextIndicator1(i).ZOrder 0 ' on top of all
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub DisplayScrollBar()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim mGroupNumber As Integer
    Dim nTimeExpectedStream As Long
    Dim nTime As Long
    
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    
    If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
        ' Midi format 0
        If MainStreamNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamNumber
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf HScrollPlayerTime.Tag > Trim$(str$(time - 2 / 86400)) Then ' still scrolling two sec
            Else
                'nTime = TimeExpectedMessage ' alternative, to scroll with original message time
                'nTime = TimeExpectedMessageRelToTempo ' alternative, to scroll with tempo
                nTime = MIDIOutput1.StreamTimeCurrent ' alternative, to scroll with stream time
                'nTime = TimeExpectedMessageRelToOpen ' not applicable, absolute time
                'nTime = TimeActualMessageRelToOpen ' not applicable, absolute time
                
                'HScrollPlayerTime.Max = ... already determined when queue
                HScrollPlayerTime.Tag = "1" ' programmatic change
                HScrollPlayerTime.value = CInt(nTime / 1000)
                HScrollPlayerTime.Tag = "" ' not needed anymore in case change() not run
                LabelQueueTime.Caption = Trim$(str$(RoundVB5(CDbl(nTime) / 1000#, 1)))
            
                ' WARNING,
                ' No way to cancel if scrolling, so will change again.
            End If
        End If
    
    Else
        ' Midi format 1
        mGroupNumber = UBound(MainStreamGroup, 1) ' last is master track
        'For mGroupNumber = 1 To UBound(MainStreamGroup, 1) ' not needed
        If mGroupNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf HScrollPlayerTime.Tag > Trim$(str$(time - 2 / 86400)) Then ' still scrolling two sec
            Else
                'nTime = TimeExpectedMessage ' alternative, to scroll with original message time
                'nTime = TimeExpectedMessageRelToTempo ' alternative, to scroll with tempo
                nTime = MIDIOutput1.StreamTimeCurrent ' alternative, to scroll with stream time
                'nTime = TimeExpectedMessageRelToOpen ' alternative, to scroll with expected absolute time
                'nTime = TimeActualMessageRelToOpen ' alternative, to scroll with actual absolute time
                
                'HScrollPlayerTime.Max = ... already determined when queue
                HScrollPlayerTime.Tag = "1" ' programmatic change
                HScrollPlayerTime.value = CInt(nTime / 1000)
                HScrollPlayerTime.Tag = "" ' not needed anymore in case change() not run
                LabelQueueTime.Caption = Trim$(str$(RoundVB5(CDbl(nTime) / 1000#, 1)))
            
                ' WARNING,
                ' No way to cancel if scrolling, so will change again.
            End If
        End If
    End If

ExitSection:

    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub ClearScrollBar()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    HScrollPlayerTime.Tag = "1" ' programmatic change
    HScrollPlayerTime.value = 0
    HScrollPlayerTime.Tag = "" ' not needed anymore, in case change() not run
    LabelQueueTime.Caption = Trim$(str$(HScrollPlayerTime.value))
    TimeExpectedMessage = 0
    TimeExpectedMessageRelToTempo = 0
    TimeExpectedMessageRelToOpen = 0
    TimeActualMessageRelToOpen = 0
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub HScrollPlayerTime_Change()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim isProgrammaticChange As Boolean
    isProgrammaticChange = False
    
    ' Distinguish between the user or the program changing the value
    If HScrollPlayerTime.Tag = "1" Then isProgrammaticChange = True
    
    If isProgrammaticChange = True Then
        ' programmatically changed
        ' do nothing
    
    Else
        ' Refresh properly in case mouseup and lostfocus events failed.
        Call CommandSetFocus_Click ' setfocus fails depending on doevents in rest of program.
        HScrollPlayerTime.Enabled = False: HScrollPlayerTime.Enabled = True ' in case setfocus fails

        ' WARNING,
        ' Setfocus may cause the scrollbar to fail depending
        ' on doevents in rest of program. But this is not a problem
        ' in all applications. Temporary solution is to use the
        ' SetFocus and Enabled properties to refresh the scrollbar.
        ' E.g.
        ' Doevents in WaitCloseStream() can screw up the scrollbar.

        If CheckMidiOutFilterFF.value = 0 Then
            Call ScrollBarPlayerTime_Forward0Common
        Else
            Call ScrollBarPlayerTime_Forward1Common
        End If
    
        ' FF performance from fast to slow, hard to easy, amount of messages:
        ' 1. FF by time and not send any messages ' fastest, easiest, not practical
        ' 2. FF by message and not send any messages ' fastest, easiest, not practical
        ' 3. FF by time and send only important messages ' faster than option 5, hardest to implement, most practical
        ' 4. FF by message and send only important messages ' faster than option 5, hardest to implement, most practical
        ' 5. FF by time, send all messages, and filter late notes ' faster than option 7, too many messages
        ' 6. FF by message, send all messages, and filter late notes ' faster than option 7, too many messages
        ' 7. Without any filtering, then option 5 and 6 are the slowest.
    End If
    
    LabelQueueTime.Caption = Trim$(str$(HScrollPlayerTime.value))
    HScrollPlayerTime.Tag = "" ' not needed anymore
    'If isProgrammaticChange = True Then ... but change() not always run
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub HScrollPlayerTime_Scroll()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim backuptimerprogressbar As Boolean
    
    ' Do not allow timer to interfere while scrolling.
    ' However, scroll is modeless so exits before done anyway.
    backuptimerprogressbar = TimerProgressBar.Enabled
    TimerProgressBar.Enabled = False
    
    ' Preview until done scrolling.
    'HScrollPlayerTime.Tag = ...  not applicable because modeless
    LabelQueueTime.Caption = Trim$(str$(HScrollPlayerTime.value))
    HScrollPlayerTime.Tag = Trim$(str$(time)) ' track time

    TimerProgressBar.Enabled = backuptimerprogressbar ' restore
    
    ' A timer can detect if have not scrolled recently.
    ' Not very realistic, but need to do something.
    'TimerProgressBar.Enabled = True

    ' WARNING,
    ' VB Scrollbars are evil.

    ' WARNING,
    ' While scrolling, the value may alternate between the
    ' scroll position and position changed programmatically.
    ' The scroll value still remains in memory though
    ' which is confusing.

    ' WARNING,
    ' Scrolling is modeless, difficult to detect completion
    ' and not multithread-safe. It was only designed to for
    ' user to change. Without detection, programmatically
    ' changing the value would interfere with it, when the user
    ' is supposed to have priority first. Usually triggers
    ' change() or validate() when mouseup() or lostfocus(),
    ' but not always. E.g. if scroll to back original value.
    
    ' WARNING,
    ' Scrolling may be affected by any doevents that run.
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub HScrollPlayerTime_Validate(Cancel As Boolean)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Refresh properly in case mouseup and lostfocus events failed.
    Call CommandSetFocus_Click ' setfocus fails depending on doevents in rest of program.
    HScrollPlayerTime.Enabled = False: HScrollPlayerTime.Enabled = True ' in case setfocus fails
    
    ' In case change() fails to run when scroll to original value.
    HScrollPlayerTime.Tag = "" ' not needed anymore
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub ScrollBarPlayerTime_Forward0Common()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close
    
    Dim isStarted As Boolean
    Dim mGroupNumber As Integer
    Dim nSeekFromTime As Long
    Dim nSeekToTime As Long
    Dim nSeekFromMessage As Long
    Dim nSeekToMessage As Long

    Dim backuplabel As String
    Const MB_WAITTEXT = "(wait)"

    isStarted = False
    
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    
    gisCurrentFF = True ' prevent multithreading issues caused by doevents
    
    ' Notify in progress
    ' (not applicable because still in progress after procedure)
    'backuplabel = LabelQueueTime.Caption
    'LabelQueueTime.Caption = MB_WAITTEXT
    'Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
    
    If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
        ' Midi format 0
        If MainStreamNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamNumber
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(0) <> MIDISTATE_STARTED Then ' not started
            Else
                isStarted = True
            End If
    
            'If CheckMidiOutFilterFF.Value = 0 Then
                
            ' Fast forward option #5
            ' Set next message based on approximate time by shifting current time.
            ' (Streamtimecurrent is SEEK FROM position, HScroll is SEEK TO position)
            MIDIOutput1.ActionStream = MIDIOUT_PAUSE
            Call StopStuckNote
            nSeekFromTime = MIDIOutput1.StreamTimeCurrent
            nSeekToTime = HScrollPlayerTime.value * MB_HSCROLLTIMESCALEOFFSET
            'nSeekFromMessage = MIDIOutput1.StreamMessageLBound ' alternative, more accurate
            'nSeekToMessage = HScrollPlayerMessage.Value * MB_HSCROLLMESSAGESCALEOFFSET ' alternative, more accurate
            If nSeekToTime >= nSeekFromTime Then
                ' fast forward, restore to prepare scan from previous message
                ' also continues from any previous fast forward calculations
                'nSeekFromTime = nSeekFromTime ' unchanged, shift start time instead
                'MIDIOutput1.StreamTimeCurrent = nSeekFromTime ' not needed
            Else
                ' rewind, restore to prepare scan from beginning
                ' also ignores any previous fast forward calculations
                nSeekFromTime = 0
                MIDIOutput1.StreamTimeCurrent = nSeekFromTime
            End If
            'MIDIOutput1.StreamTimeStartRelativeToOpen = ... adjust start time
            ' Fast forward calculations are continued later to
            ' adjust the start time when starting playback.
        End If

    Else
        ' Midi format 1
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            ' Stop all streams quickly
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(0) <> MIDISTATE_STARTED Then ' not started
            Else
                isStarted = True
            End If

            'If CheckMidiOutFilterFF.Value = 0 Then

            ' Fast forward option #5
            ' Set next message based on approximate time by shifting current time.
            ' (Streamtimecurrent is SEEK FROM position, HScroll is SEEK TO position)
            MIDIOutput1.ActionStream = MIDIOUT_PAUSE
            Call StopStuckNote
        Next mGroupNumber
        
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            ' Process fast forward
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            nSeekFromTime = MIDIOutput1.StreamTimeCurrent
            nSeekToTime = HScrollPlayerTime.value * MB_HSCROLLTIMESCALEOFFSET
            'nSeekFromMessage = MIDIOutput1.StreamMessageLBound ' alternative, more accurate
            'nSeekToMessage = HScrollPlayerMessage.Value * MB_HSCROLLMESSAGESCALEOFFSET ' alternative, more accurate
            If nSeekToTime >= nSeekFromTime Then
                ' fast forward, restore to prepare scan from previous message
                ' also continues from any previous fast forward calculations
                'nSeekFromTime = nSeekFromTime ' unchanged, shift start time instead
                'MIDIOutput1.StreamTimeCurrent = nSeekFromTime ' not needed
            Else
                ' rewind, restore to prepare scan from beginning
                ' also ignores any previous fast forward calculations
                nSeekFromTime = 0
                MIDIOutput1.StreamTimeCurrent = nSeekFromTime
            End If
            'MIDIOutput1.StreamTimeStartRelativeToOpen = ... adjust start time
            ' Fast forward calculations are continued later to
            ' adjust the start time when starting playback.
        Next mGroupNumber
    End If
    
    ' Keep started since paused.
    ' Approximate time by shifting start time.
    If isStarted = True Then StartPlay

ExitSection:
    
    gisCurrentFF = False ' not needed anymore

    'If LabelQueueTime.Caption = MB_WAITTEXT Then _
    ' LabelQueueTime.Caption = backuplabel ' restore, only if not already changed
    
    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub ScrollBarPlayerTime_Forward1Common()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close
    
    Dim isStarted As Boolean
    Dim mGroupNumber As Integer
    Dim nSeekFromTime As Long
    Dim nSeekToTime As Long
    Dim nSeekFromMessage As Long
    Dim nSeekToMessage As Long
    Dim longtemp As Long

    Dim cStreamname2 As String
    Dim mStreamNumber2 As Integer
    Dim backuplabel As String
    Const MB_WAITTEXT = "(wait)"

    isStarted = False
    
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    
    ' Notify in progress
    backuplabel = LabelQueueTime.Caption
    LabelQueueTime.Caption = MB_WAITTEXT
    Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents

    ' Prepare stream for skipped range of messages
    cStreamname2 = MB_STREAMNAME_FF
    mStreamNumber2 = 0
    Call OpenQueueStream(mStreamNumber2, cStreamname2, MIDIOutput1)
    
    If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
        ' Midi format 0
        If MainStreamNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamNumber
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(0) <> MIDISTATE_STARTED Then ' not started
            Else
                isStarted = True
            End If
    
            'If CheckMidiOutFilterFF.Value = 0 Then
                
            ' Fast forward option #3
            ' Set next message and send only important messages that were skipped.
            ' (Streamtimecurrent is SEEK FROM position, HScroll is SEEK TO position)
            MIDIOutput1.StreamTimeCurrent = HScrollPlayerTime.value * MB_HSCROLLTIMESCALEOFFSET
            'MIDIOutput1.StreamMessageLBound = HScrollPlayerMessage.Value * MB_HSCROLLMESSAGESCALEOFFSET ' alternative, more accurate
            'MIDIOutput1.ActionStream = MIDIOUT_PAUSE ' already done by StreamTimeCurrent
            'MIDIOutput1.StreamTimeCurrentPrevious = MIDIOutput1.StreamTimeCurrent ' already done by StreamTimeCurrent
            Call StopStuckNote
            nSeekFromTime = MIDIOutput1.StreamTimeCurrentPrevious
            nSeekToTime = MIDIOutput1.StreamTimeCurrent
            nSeekFromMessage = MIDIOutput1.StreamMessageLBoundPrevious ' alternative, more accurate
            nSeekToMessage = MIDIOutput1.StreamMessageLBound ' alternative, more accurate
            If nSeekToTime >= nSeekFromTime Then
                ' fast forward, restore to prepare scan from previous message
                ' also continues from any previous fast forward calculations
                'nSeekFromTime = nSeekFromTime ' unchanged, shift start time instead
                'MIDIOutput1.StreamTimeCurrent = nSeekFromTime ' not needed
                'nSeekFromMessage = nSeeToMessage ' unchanged, shift start time instead
                'MIDIOutput1.StreamMessageLBound = nSeekFromMessage = ' not needed

                ' Copy messages within skipped range in a new stream
                ' (help speed up by filtering messages and keep important messages)
                Call CommandStreamScanCopyMessage_ClickCommon(MIDIOutput1.StreamNumber, mStreamNumber2 _
                 , nSeekFromMessage, nSeekToMessage, isSortOutOfOrder:=True _
                 , isFilterNonEnabled:=True, isFilterNotes:=True)
                Call FilterDuplicateMessages(mStreamNumber2)
            
            Else
                ' rewind, restore to prepare scan from beginning
                ' also ignores any previous fast forward calculations
                nSeekFromTime = 0
                'MIDIOutput1.StreamTimeCurrent = nSeekFromTime ' not applicable
                nSeekFromMessage = 0
                'MIDIOutput1.StreamMessageLBound = nSeekFromMessage = ' not applicable
            
                ' Copy messages within skipped range in a new stream
                ' (help speed up by filtering messages and keep important messages)
                Call CommandStreamScanCopyMessage_ClickCommon(MIDIOutput1.StreamNumber, mStreamNumber2 _
                 , nSeekFromMessage, nSeekToMessage, isSortOutOfOrder:=True _
                 , isFilterNonEnabled:=True, isFilterNotes:=True)
                Call FilterDuplicateMessages(mStreamNumber2)
            End If
           
            ' Send messages
            '{
                MIDIOutput1.StreamNumber = mStreamNumber2
                
                ' Sort stream, by time, message number, channel, message+data1
                ' (already sorted by channel if midi 1 file format)
                ' no
                     
                ' Start queue stream and inherit properties, or let calling program do it
                MIDIOutput1.StreamAutoClose = True ' not need anymore
                MIDIOutput1.FilterLateEventStreamMax = True ' may filter notes
                Call CheckMidiOutFilterLateEventAllMax_Click ' may filter notes
                'MIDIOutput1.StreamTimeStartRelativeToOpen = MB_VERY_OLD_TIME ' very old time to send fast ignoring message time
                MIDIOutput1.StreamTempoRate = 1 ' send fast ignoring message time
                'MIDIOutput1.StreamTimeStartPending = True ' as soon as possible, hides the time (default)
                MIDIOutput1.ActionStream = MIDIOUT_START
                    
                ' Wait until all related streams be empty and autoclose
                Call WaitStopStream(mStreamNumber2, cStreamname2, MIDIOutput1)
            '}
            
            'MIDIOutput1.StreamTimeStartRelativeToOpen = ... adjust start time
            ' Fast forward calculations are continued later to
            ' adjust the start time when starting playback.
        End If

    Else
        ' Midi format 1
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            ' Stop all streams quickly
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(0) <> MIDISTATE_STARTED Then ' not started
            Else
                isStarted = True
            End If
        
            'If CheckMidiOutFilterFF.Value = 0 Then
                
            ' Fast forward option #3
            ' Set next message and send only important messages that were skipped.
            ' (Streamtimecurrent is SEEK FROM position, HScroll is SEEK TO position)
            MIDIOutput1.StreamTimeCurrent = HScrollPlayerTime.value * MB_HSCROLLTIMESCALEOFFSET
            'MIDIOutput1.StreamMessageLBound = HScrollPlayerMessage.Value * MB_HSCROLLMESSAGESCALEOFFSET ' alternative, more accurate
            'MIDIOutput1.ActionStream = MIDIOUT_PAUSE ' already done by StreamTimeCurrent
            'MIDIOutput1.StreamTimeCurrentPrevious = MIDIOutput1.StreamTimeCurrent ' already done by StreamTimeCurrent
            Call StopStuckNote
        Next mGroupNumber
        
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            ' Process fast forward
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            nSeekFromTime = MIDIOutput1.StreamTimeCurrentPrevious
            nSeekToTime = MIDIOutput1.StreamTimeCurrent
            nSeekFromMessage = MIDIOutput1.StreamMessageLBoundPrevious ' alternative, more accurate
            nSeekToMessage = MIDIOutput1.StreamMessageLBound ' alternative, more accurate
            If nSeekToTime >= nSeekFromTime Then
                ' fast forward, restore to prepare scan from previous message
                ' also continues from any previous fast forward calculations
                'nSeekFromTime = nSeekFromTime ' unchanged, shift start time instead
                'MIDIOutput1.StreamTimeCurrent = nSeekFromTime ' not needed
                'nSeekFromMessage = nSeeToMessage ' unchanged, shift start time instead
                'MIDIOutput1.StreamMessageLBound = nSeekFromMessage = ' not needed

                ' Copy messages within skipped range in a new stream
                ' (help speed up by filtering messages and keep important messages)
                Call CommandStreamScanCopyMessage_ClickCommon(MIDIOutput1.StreamNumber, mStreamNumber2 _
                 , nSeekFromMessage, nSeekToMessage, isSortOutOfOrder:=True _
                 , isFilterNonEnabled:=True, isFilterNotes:=True)
                'Call FilterDuplicateMessages(mStreamNumber2) ' do after all streams
            
            Else
                ' rewind, restore to prepare scan from beginning
                ' also ignores any previous fast forward calculations
                nSeekFromTime = 0
                'MIDIOutput1.StreamTimeCurrent = nSeekFromTime ' not applicable
                nSeekFromMessage = 0
                'MIDIOutput1.StreamMessageLBound = nSeekFromMessage = ' not applicable
            
                ' Copy messages within skipped range in a new stream
                ' (help speed up by filtering messages and keep important messages)
                Call CommandStreamScanCopyMessage_ClickCommon(MIDIOutput1.StreamNumber, mStreamNumber2 _
                 , nSeekFromMessage, nSeekToMessage, isSortOutOfOrder:=True _
                 , isFilterNonEnabled:=True, isFilterNotes:=True)
                'Call FilterDuplicateMessages(mStreamNumber2) ' do after all streams
            End If
        Next mGroupNumber
        
        Call FilterDuplicateMessages(mStreamNumber2)
        
        ' Send messages
        '{
            MIDIOutput1.StreamNumber = mStreamNumber2
            
            ' Sort stream, by time, message number, channel, message+data1
            ' (already sorted by channel if midi 1 file format)
            ' no
            
            ' Start queue stream and inherit properties, or let calling program do it
            MIDIOutput1.StreamAutoClose = True ' not need anymore
            MIDIOutput1.FilterLateEventStreamMax = True ' may filter notes
            Call CheckMidiOutFilterLateEventAllMax_Click ' may filter notes
            'MIDIOutput1.StreamTimeStartRelativeToOpen = MB_VERY_OLD_TIME ' very old time to send fast ignoring message time
            MIDIOutput1.StreamTempoRate = 1 ' send fast ignoring message time
            'MIDIOutput1.StreamTimeStartPending = True ' as soon as possible, hides the time (default)
            MIDIOutput1.ActionStream = MIDIOUT_START
                
            ' Wait until all related streams be empty and autoclose
            Call WaitCloseStream(mStreamNumber2, cStreamname2, MIDIOutput1)
            
#If 1 = 0 Then ' comment out to enable test
            ' check status
            Debug.Print MIDIOutput1.StreamNumber _
             , MIDIOutput1.StateStreamEx(0) _
             , MIDIOutput1.StateStreamEx(mStreamNumber2) _
             , mStreamNumber2 _
             , Rnd()
#End If
            
        '}
        
        'MIDIOutput1.StreamTimeStartRelativeToOpen = ... adjust start time
        ' Fast forward calculations are continued later to
        ' adjust the start time when starting playback.
    End If
    
    ' Keep started since paused.
    ' Approximate time by shifting start time.
    If isStarted = True Then StartPlay

ExitSection:

    If LabelQueueTime.Caption = MB_WAITTEXT Then _
     LabelQueueTime.Caption = backuplabel ' restore, only if not already changed
    
    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub ClearPlayerButtons()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    CmdPlay.Enabled = False
    CmdPause.Enabled = False
    CmdStop.Enabled = False
    CheckAutoReplay.value = 0: Call CheckAutoReplay_Click
    CheckAutoStop.value = 1: Call CheckAutoStop_Click
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub DisplayPlayerButtons()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    CmdPlay.Enabled = True
    CmdPause.Enabled = True
    CmdStop.Enabled = True
    CheckAutoReplay.value = 0: Call CheckAutoReplay_Click
    CheckAutoStop.value = 1: Call CheckAutoStop_Click
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub ClearPlayer()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Call MIDIOutput1_Unload
    MainStreamNumber = 0 ' initially no streams
    ReDim MainStreamGroup(0) ' initially no streams
    MainMidifile = "" ' no file open
    Me.Caption = MainMidifile
    LabelPlayerFile.Caption = ""
    LabelNumberOfTracks.Caption = ""
    Call ClearTrackBars
    Call ClearScrollBar
    Call ClearPlayerButtons
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub MIDIOutput1_Initialize()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Open selected device
    If MIDIOutput1.State <> MIDISTATE_CLOSED Then Call MIDIOutput1_Unload
    MIDIOutput1.DeviceID = OutputDevCombo.ListIndex - 1
    MIDIOutput1.action = MIDIOUT_OPEN
    'If MIDIOutput1.ErrorCode <> 0 Then ... (optional)
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub MIDIOutput1_Unload()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Open selected device
    If MIDIOutput1.State <> MIDISTATE_CLOSED Then
        StopPlay ' clear any stuck notes
        MIDIOutput1.action = MIDIOUT_CLOSE
    End If
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub MIDIOutput1_Error(ErrorCode As Integer, ErrorMessage As String)
    Err.Raise 1, , "PROGRAM ERROR 27565, wrong errorscheme"
    'If ErrorCode = 0 Or ErrorCode = 4 Or ErrorCode = 8 Then
    ' (not using it anymore)
End Sub

Private Sub MIDIOutput1_ErrorEx(Number As Integer, Description As String, Source As String)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    If Number = 0 Or Number = 1004 Or Number = 1008 Then
        ' Error is known, can be handled and continued.
        Call ShowMidiError(Number, Description, Source _
         , MIDIOutput1.ErrorCount, False)
    Else
        ' Other error stops.
        Call ShowMidiError(Number, Description, Source _
         , MIDIOutput1.ErrorCount, True)
    End If
    MIDIOutput1.ErrorCode = 0 ' not needed anymore
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub MIDIOutput1_MessageSent(MessageTag As Long)
    ' (moved to StreamSend())
End Sub

Private Sub MIDIOutput1_QueueEmpty()
    ' (moved to StreamEmpty())
End Sub

Private Sub MIDIOutput1_StreamEmpty(StreamNumber As Integer, StateStream As Integer, StreamName As String)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim mGroupNumber As Integer
    Dim isStop As Boolean
    
#If 1 = 0 Then ' comment out to enable test
    Debug.Print "-----------------------------"
    Debug.Print StreamNumber, StateStream, StreamName
#End If
    
    ' Distinuigh between FF stream, midi 0 stream amd midi 1 streams
    If StreamName = MB_STREAMNAME_FF Then
        ' autocloses
    
    ElseIf StreamName = MB_STREAMNAME_1 Then
        If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
            If MainStreamNumber <> 0 Then
                If StreamNumber = MainStreamNumber Then
                    Call CmdStop_Click ' main stream stopped
                End If
            End If
    
        Else
            ' Midi format 1
            ' Only stop once all streams empty.
            isStop = True ' assume all streams stopped
            For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
                If MIDIOutput1.StateStreamEx(MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)) = MIDISTATE_STARTED Then
                    isStop = False ' at least one stream still processing
                End If
            Next mGroupNumber
            If isStop = True Then Call CmdStop_Click
        End If
    End If
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub MIDIOutput1_StreamSend()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    '-------------------------------------------------------------------
    ' Collect information for indicators and progress.
    
    ' WARNING,
    ' Operations that do not need to be in real time
    ' could be located in other timers to keep procedure
    ' as small and fast as possible.

    ' WARNING,
    ' Operations that need to access public storage
    ' in the midi controls must be located in other
    ' timers because not multithread-safe here.
    '-------------------------------------------------------------------
    
    Dim MessageTag(TOTAL_MIDI_CHANNELS) As Long
    Dim TrackNumber As Integer
    Dim Intensity As Integer
    Dim channel As Integer
    Dim m As Integer
    Dim nTime As Long
    Dim nTimeRelToTempo As Long
    
    '
    ' Get last messagetag in each channel
    '
    nTime = MIDIOutput1.SendTime ' initialize
    nTimeRelToTempo = MIDIOutput1.SendTimeRelativeToTempo ' initialize
    Do While MIDIOutput1.MessageCount > 0 ' (optional)
        If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
        
        channel = (MIDIOutput1.SendMessage And &HF) + 1
        If MIDIOutput1.SendMessageTag = 0 Then ' no tag
        ElseIf MIDIOutput1.SendMessageTag < MessageTag(channel) Then ' not a peak
        Else
            ' Overwrite and discard old message tags, if any
            MessageTag(channel) = MIDIOutput1.SendMessageTag
        End If
        
        ' Track oldest message to reflect in scroll bar
        ' (optional)
        If MIDIOutput1.SendTime < nTime Then nTime = MIDIOutput1.SendTime
        If MIDIOutput1.SendTimeRelativeToTempo < nTimeRelToTempo Then nTimeRelToTempo = MIDIOutput1.SendTimeRelativeToTempo
        
        MIDIOutput1.ActionStream = MIDIOUT_REMOVE
    
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' Release resources for background processing
        ' back to object
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' (not needed in upgrade)
        ' (not multithread-safe to use DoEvents here)
        'Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
    
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' Release resources for background processing
        ' by terminating immediately
        ''''''''''''''''''''''''''''''''''''''''''''''
        'Exit Do ' never terminate early, must process all, at most 100 per event
    Loop
    
    TimeExpectedMessage = nTime
    TimeExpectedMessageRelToTempo = nTimeRelToTempo
    'TimeExpectedMessageRelToOpen = MIDIOutput1.SendTimeCurrentRelativeToOpen ' not needed
    'TimeActualMessageRelToOpen = MIDIOutput1.SendTimeActualRelativeToOpen ' not needed
    'TimeActualMessageRelToOpen = MIDIOutput1.TimeRelativeToOpen ' alternative, but slow
    
    '
    ' Process last messagetag in each channel
    '
    For channel = 1 To TOTAL_MIDI_CHANNELS
        If MessageTag(channel) = 0 Then ' no tag
        ElseIf (MessageTag(channel) < 0) Or (MessageTag(channel) >= 32000) Then ' not applicable
        Else
            Intensity = MessageTag(channel) Mod 1000
            TrackNumber = Int(MessageTag(channel) / 1000)
            TrackNumber = TrackNumber + 1 ' restore to 1-based scale to match other arrays
        
            ' Track note in progress bar
            If (TrackNumber > 0) And (Intensity > VIndicator1(TrackNumber - TrackOffset).value) Then
            'If (Intensity > 0) And (TrackNumber > 0) And (Intensity > Vindicator1(TrackNumber - TrackOffset).Value) Then ' ignore zero velocity
                VIndicator1(TrackNumber - TrackOffset).value = Intensity
            End If
            
            ' Track note in progress bar peak
            If CheckPeakHold.value = 0 Then ' not needed
            ElseIf (TrackNumber > 0) And (Intensity > VIndicatorPeak(TrackNumber - TrackOffset).value) Then
                ' similar to other bar
                VIndicatorPeak(TrackNumber - TrackOffset).value = Intensity
                CommandIndicatorPeak(TrackNumber - TrackOffset).Visible = True
                CommandIndicatorPeak(TrackNumber - TrackOffset).Tag = "" ' reset for new note
            End If
        End If
    Next channel

    ' Update queuetime
    'MIDIOutput1.StreamNumber = MIDIOutput1.SendStreamNumber
    'LabelQueueTime.Caption = Trim$(Str$(RoundVB5(CDbl(MIDIOutput1.StreamTimeCurrent) / 1000#, 1))) & " seconds"
    ' (not multithread-safe here)

    ' Trigger other extensive operations.
    'TimerProgressBar.Enabled = True ' already set

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub OptionOpen_Click(index As Integer)
    '
End Sub

Private Sub OutputDevCombo_Click()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Prepare form for close
    ' assuming still old port shown in combobox until changed
    
    If 1 = 0 Then
        MIDIOutput1.ReOpen ' not yet implemented
        
    ElseIf 1 = 0 Then
        ' Prepare midi file and midi data
        Dim isOpen As Boolean
        If isOpen = True Then StartPlay
    
    Else
        Call ClearPlayer
    End If
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
   
Private Sub QueueSong_ByMidi0Pieces()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    '-------------------------------------------------------------------
    ' Queue Song
    ' - scan for global messages, like tempo
    ' - queue pieces of each track or groups of messages
    '   which are close in time to speed up queuesort, manualsort, autosort
    ' - insert and queue global messages, as needed
    ' - queue messages with MessagePointer
    ' - queue message to a stream
    '-------------------------------------------------------------------
    
    Debug.Print "Verifying: Queue midi 0 by pieces, " _
     ; IIf(CheckManualSort.value = 0, "", "sort manually")
    
    Dim backupscreenmousepointer As Integer
    Dim backupstreammessagenumbermax As Long
    Dim mTrackPhysical As Integer
    Dim mTrackLogical As Integer
    Dim isTrackMute As Boolean
    Dim isTrackDone As Boolean
    Dim mTrackLoadComplete As Integer
    Dim mTrackLoadTotal As Integer
    Dim nTrackMessageTimeCurrent As Long
    Dim nTrackMessageTimeIncrement As Long
    Dim nTrackMessageTimeTo As Long
    Dim nTrackMessageTimeNext As Long
    Dim nTrackMessageCurrent As Long
    Dim nTrackMessageIncrement As Long
    Dim nTrackMessageTo As Long
    Dim isTrackMessageEnd As Boolean
    Dim arTrackDone() As Boolean
    Dim arTrackMessageTimeCurrent() As Long
    Dim arTrackMessageTimeNext() As Long
    Dim arTrackMessageCurrent() As Long
    Dim m As Long
    Dim i As Long
    Dim nMessageCount As Long
    Dim nMessageTotal As Long
    Dim mR As Long
    Dim mC As Long
    
    Dim MIDIOutput1_MP(0 To MIDIMP_UBOUND) As Long ' always start from zero
    Dim nMP As Long
    Dim tempmessage As Integer
    Dim tempdata1 As Integer
    Dim tempdata2 As Integer
    Dim temptime As Long
    Dim tempmessagetag As Long
    Dim tempmessagestate As Integer
    Dim templogonly As Boolean
    
    Dim mMsgFF81TempoCount As Integer
    Dim mMsgFF88TPQCount As Integer
    Dim mMsgFF81TempoCountMax As Integer
    Dim mMsgFF88TPQCountMax As Integer
    Dim arMsgFF81Tempo() As Long
    Dim arMsgFF88TPQ() As Long
    Dim arMsgFF81TempoCount() As Integer
    Dim arMsgFF88TPQCount() As Integer
    Const MB_DIMENSION1UBOUND = 3
    Const MB_TICK = 1
    Const MB_VALUE = 2
    Const MB_TICKNEXT = 3

    Dim backuptempo As Long
    Dim backupticksperquarternote As Integer
    Dim backupnumerator As Integer
    Dim backupdenominator As Integer
    Dim dTicksPerMillisecond As Double
    Dim nTicksBetweenEvents As Long
    Dim nTicksRemaining As Long
    Dim nMillisecondsBetweenEvents As Long
    Dim nStreamTimeCurrent As Long
    Dim nStreamTicksCurrent As Long
    Dim isTrackTicks As Boolean
    Dim nStreamTimeStart As Long
    Dim isGlobal As Boolean
    Dim isMsgFF81TempoChange As Boolean
    Dim isMsgFF88TPQChange As Boolean
    Dim isSortOutOfOrder As Boolean
    Dim arStreamTicksCurrent() As Long
    Dim arStreamTimeCurrent() As Long
    Dim arStreamTimeStart() As Long
    Dim arTicksPerMillisecond() As Double
    Dim arTempoPrevious() As Long
    
    Dim nStartRelativeToStream As Long
    Dim nCurrentRelativeToStream As Long
    Dim dTimeDifferenceOld As Double
    Dim dTimeDifference100 As Double
    Dim dTimeDifferenceNew As Double
    Dim dTempo As Double
    Dim nTempoCurrent As Long
    Dim nTempoPrevious As Long
    Dim isProcessTempo As Boolean

    If (MIDIFile1.Filename = "") Then GoTo ExitEnd

    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    backupscreenmousepointer = Screen.MousePointer
    Screen.MousePointer = 11
    Me.Tag = Me.Caption ' backup

    ' Pick one stream number the whole form can use.
    Call OpenQueueStream(MainStreamNumber, MB_STREAMNAME_1, MIDIOutput1)
    MIDIOutput1.StreamNumber = MainStreamNumber
    
    ' Clear any data if stream not new
    MIDIOutput1.ActionStream = MIDIOUT_RESET

    ' Prevent queuesort temporarily for speed
    ' in case too many messages in a midi file
    ' format 1 so not sort on every message.
    ' Assume will sort manually later.
    If CheckManualSort.value = 1 Then MIDIOutput1.StreamMessageSortOutOfOrder = True
    
    ' Get statistics
    nMessageTotal = 0
    backupstreammessagenumbermax = MIDIOutput1.StreamMessageNumberMax
    For m = 1 To MIDIFile1.NumberOfTracks
        MIDIFile1.TrackNumber = m
        nMessageTotal = nMessageTotal + MIDIFile1.MessageCount
    Next m
    If MIDIFile1.NumberOfTracks = 1 Then
        ' Midi 0 file format
        ' All messages merged and sorted already.
    Else
        ' Midi 1 file format
        ' First track is usually global midi info, but not guarenteed.
        ' Second track is first track with notes.
        ' Etc.
    End If
    
    Me.Caption = "MFPlayer Example - Loading - " _
     & Trim$(str$(Int(100 * nMessageCount / nMessageTotal))) & "%"
    
    ' Get global tags for reference.
    ' Assume in one track if midi file format 0.
    ' Assume in first track if midi file format 1. May be in others but not standard.
    '{
        ' Get tempo info
        MIDIFile1.TrackNumber = 1
        MIDIFile1.MessageNumber = 0
        backuptempo = MIDIFile1.Tempo
        backupticksperquarternote = MIDIFile1.TicksPerQuarterNote
        backupnumerator = MIDIFile1.Numerator
        backupdenominator = 2 ^ MIDIFile1.Denominator
        If backuptempo = 0 Then backuptempo = 600000 ' assume 100 beats per minute (tempo/2 = beats*2)
        If backupticksperquarternote = 0 Then backupticksperquarternote = 480 ' assume 100 beats per minute
        If backupnumerator = 0 Then backupnumerator = 4 ' assume time signature 4/4
        If backupdenominator = 0 Then backupdenominator = 4
        dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000#
        
        ReDim arMsgFF81Tempo(MB_DIMENSION1UBOUND, 0 To 1000)
        ReDim arMsgFF88TPQ(MB_DIMENSION1UBOUND, 0 To 1000)
        
        mMsgFF81TempoCount = 0
        arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) = 0 ' tick zero
        arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount) = backuptempo
        arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) = 0 ' not yet
        
        mMsgFF88TPQCount = 0
        arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) = 0 ' tick zero
        arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount) = backupticksperquarternote
        arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) = 0 ' not yet
        
        nMessageCount = 0
        isSortOutOfOrder = False
        For mTrackPhysical = 1 To MIDIFile1.NumberOfTracks ' 1-based scale
            MIDIFile1.TrackNumber = mTrackPhysical ' 1-based scale (first is global, second is track one)
            mTrackLogical = mTrackPhysical - 1 ' 0-based scale (zero is global, first is track one)
        
            nStreamTicksCurrent = 0
            For m = 1 To MIDIFile1.MessageCount ' 1-based scale
                If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                
                MIDIFile1.MessageNumber = m ' 1-based scale
        
                If Int(m / 1000) = m / 1000 And nMessageTotal <> 0 Then
                    Me.Caption = "MFPlayer Example - Loading - " _
                     & IIf(nMessageCount < nMessageTotal, "0", "") _
                     & Trim$(str$(Int(100 * nMessageCount / nMessageTotal) / 100)) & "%"
                    ' fraction to show preloading
                    Call DoEventsOnce(True): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
                End If
                
                nTicksBetweenEvents = MIDIFile1.time
                nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                
                tempmessage = MIDIFile1.Message
                If tempmessage <> META Then 'ignore
                ElseIf MIDIFile1.Data1 = META_TEMPO Then ' tempo
                    mMsgFF81TempoCount = mMsgFF81TempoCount + 1
                    If mMsgFF81TempoCount > UBound(arMsgFF81Tempo, 2) Then _
                     ReDim Preserve arMsgFF81Tempo(MB_DIMENSION1UBOUND, UBound(arMsgFF81Tempo, 2) + 100) ' more space
                    
                    arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) = nStreamTicksCurrent
                    arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount) = MIDIFile1.Tempo
                    arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) = 0 ' not yet
                    arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount - 1) = nStreamTicksCurrent ' save for reference
                
                    If arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) < arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount - 1) Then _
                     isSortOutOfOrder = True ' when message not in first track
                
                ElseIf MIDIFile1.Data1 = 88 Then ' time sig
                    mMsgFF88TPQCount = mMsgFF88TPQCount + 1
                    If mMsgFF88TPQCount > UBound(arMsgFF88TPQ, 2) Then _
                     ReDim Preserve arMsgFF88TPQ(MB_DIMENSION1UBOUND, UBound(arMsgFF88TPQ, 2) + 100) ' more space
                    
                    arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) = nStreamTicksCurrent
                    arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount) = MIDIFile1.TicksPerQuarterNote
                    arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) = 0 ' not yet
                    arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount - 1) = nStreamTicksCurrent ' save for reference
                
                    If arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) < arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount - 1) Then _
                     isSortOutOfOrder = True ' when message not in first track
                
                End If
            
                nMessageCount = nMessageCount + 1
                
                If nMessageCount = backupstreammessagenumbermax Then Exit For ' reached limit
                'If nMessageCount = MB_LONGUBOUND Then Exit For ' not accurate
                'If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' not applicable
            Next m
            
            If nMessageCount = backupstreammessagenumbermax Then Exit For ' reached limit
            'If nMessageCount = MB_LONGUBOUND Then Exit For ' not accurate
            'If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' not applicable
        Next mTrackPhysical
        
        mMsgFF81TempoCountMax = mMsgFF81TempoCount
        mMsgFF88TPQCountMax = mMsgFF88TPQCount
    
#If 1 = 0 Then ' comment out to enable test
        Debug.Print "------------------------------------------"
        Debug.Print "arMsgFF81Tempo"
        For mR = LBound(arMsgFF81Tempo, 2) To mMsgFF81TempoCountMax
         Debug.Print mR; Space$(5);
         For mC = 1 To UBound(arMsgFF81Tempo, 1)
            Debug.Print arMsgFF81Tempo(mR, mC) & Space$(5);
         Next
         Debug.Print
        Next
        Stop
#End If
    
    '}
    
    ' Get all messages
    nMessageCount = 0
    nTrackMessageTimeCurrent = 0
    nTrackMessageTimeTo = 0
    nTrackMessageTimeIncrement = 5000 ' in msec
    'nTrackMessageTimeIncrement = 100 ' too slow, initializing variables
    'nTrackMessageTimeIncrement = 100000 ' too slow, large sorting range
    ReDim arTrackDone(MIDIFile1.NumberOfTracks)
    ReDim arTrackMessageTimeNext(MIDIFile1.NumberOfTracks)
    ReDim arTrackMessageCurrent(MIDIFile1.NumberOfTracks)
    ReDim arMsgFF81TempoCount(MIDIFile1.NumberOfTracks)
    ReDim arMsgFF88TPQCount(MIDIFile1.NumberOfTracks)
    ReDim arStreamTicksCurrent(MIDIFile1.NumberOfTracks)
    ReDim arStreamTimeCurrent(MIDIFile1.NumberOfTracks)
    ReDim arStreamTimeStart(MIDIFile1.NumberOfTracks)
    ReDim arTicksPerMillisecond(MIDIFile1.NumberOfTracks)
    ReDim arTempoPrevious(MIDIFile1.NumberOfTracks)
    
    mTrackLoadTotal = MIDIFile1.NumberOfTracks ' reduce total later if tracks ignored
    Do While mTrackLoadComplete < mTrackLoadTotal
        
        ' Increment through in groups for speed when sorting.
        ' Must track by message time and not message number.
        ' Synchronize with all tracks the same.
        nTrackMessageTimeTo = nTrackMessageTimeCurrent + nTrackMessageTimeIncrement

        ' Scan remaining tracks for messages.
        For mTrackPhysical = 1 To MIDIFile1.NumberOfTracks ' 1-based scale
            MIDIFile1.TrackNumber = mTrackPhysical ' 1-based scale (first is global, second is track one)
            mTrackLogical = mTrackPhysical - 1 ' 0-based scale (zero is global, first is track one)
            isTrackMute = False
            
            ' Initialize or load from last scan.
            If arTempoPrevious(mTrackLogical) = 0 Then ' zero not applicable
                arTempoPrevious(mTrackLogical) = backuptempo ' default is backuptempo
            End If
            If arTicksPerMillisecond(mTrackLogical) = 0 Then ' zero not applicable
                arTicksPerMillisecond(mTrackLogical) = dTicksPerMillisecond ' default is dTicksPerMillisecond
            End If
            isTrackDone = arTrackDone(mTrackLogical) ' default is false
            nTrackMessageTimeNext = arTrackMessageTimeNext(mTrackLogical) ' default is 0
            nTrackMessageCurrent = arTrackMessageCurrent(mTrackLogical) ' default is 0
            nStreamTicksCurrent = arStreamTicksCurrent(mTrackLogical) ' default is 0
            nStreamTimeCurrent = arStreamTimeCurrent(mTrackLogical) ' default is 0
            nStreamTimeStart = arStreamTimeStart(mTrackLogical) ' default is 0
            dTicksPerMillisecond = arTicksPerMillisecond(mTrackLogical) ' default is dTicksPerMillisecond
            nTempoPrevious = arTempoPrevious(mTrackLogical) ' default is backuptempo
            mMsgFF81TempoCount = arMsgFF81TempoCount(mTrackLogical) ' default is 0
            mMsgFF88TPQCount = arMsgFF88TPQCount(mTrackLogical) ' default is 0

#If 1 = 0 Then ' comment out to enable test
            ' Mute some tracks (not channels).
            ' Assuming midi file format 1.
            ' Assuming track zero has all global messages.
            If mTrackLogical <> 0 _
             And mTrackLogical <> 8 _
             And mTrackLogical <> 9 _
             Then
                isTrackMute = True
                mTrackLoadTotal = 3 ' number of tracks actually processing
            End If
#End If
        
            If isTrackMute = True Then
            
            ElseIf isTrackDone = True Then
                ' Already loaded all messages from track.

            ElseIf nTrackMessageTimeTo < nTrackMessageTimeNext Then
                ' Not reach message yet in track.
            
            Else
                nTrackMessageTo = MIDIFile1.MessageCount ' assume to end
                nTrackMessageIncrement = 0
                isTrackMessageEnd = False
                
                If nTrackMessageCurrent + 1 > nTrackMessageTo Then
                    isTrackMessageEnd = True ' no more messages
                End If
                
                For m = nTrackMessageCurrent + 1 To nTrackMessageTo ' 1-based scale
                    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                    If isTrackMessageEnd = True Then Exit For
                    
                    If Int(m / 250) = m / 250 And nMessageTotal <> 0 Then
                        Me.Caption = "MFPlayer Example - Loading - " _
                         & Trim$(str$(Int(100 * nMessageCount / nMessageTotal))) & "%"
                        Call DoEventsOnce(True): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
                    End If
        
                    MIDIFile1.MessageNumber = m ' 1-based scale
                
                    ' Get next time
                    nTicksRemaining = MIDIFile1.time
                    
                    ' Insert global messages, if any
                    ' if occurs before next message.
                    Do
                        ' Assume not scan entire array for tempo since sequential and sorted.
                        isGlobal = False
                        isMsgFF81TempoChange = False
                        isMsgFF88TPQChange = False
                        If mMsgFF81TempoCount <> mMsgFF81TempoCountMax _
                         And nStreamTicksCurrent + nTicksRemaining >= arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) Then
                            isGlobal = True
                            isMsgFF81TempoChange = True
                            mMsgFF81TempoCount = mMsgFF81TempoCount + 1
                            nTicksBetweenEvents = arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) - nStreamTicksCurrent
                        
                        ElseIf mMsgFF88TPQCount <> mMsgFF88TPQCountMax _
                         And nStreamTicksCurrent + nTicksRemaining >= arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) Then
                            isGlobal = True
                            isMsgFF88TPQChange = True
                            mMsgFF88TPQCount = mMsgFF88TPQCount + 1
                            nTicksBetweenEvents = arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) - nStreamTicksCurrent
                        End If
                    
                        If isGlobal = False Then Exit Do ' none
                    
                        ' Get time
                        ' Assuming all previous ticks were for one tempo only.
                        ' Assuming shifting start time already compensated for.
                        ' Assume tracking ticks is more accurate than tracking time.
                        nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                        nStreamTimeCurrent = RoundVB5(nStreamTicksCurrent / dTicksPerMillisecond, 0) ' ticks to time
                    
                        ' Adjust for changes in tempo.
                        If isMsgFF81TempoChange = True Then
                            ' New start time.
                            ' Shift start time to compensate based on
                            ' current time and tempo rate.
                            '{
                                ' estimated message current time
                                nStartRelativeToStream = 0
                                nCurrentRelativeToStream = nStreamTimeCurrent
                                dTimeDifferenceOld = nCurrentRelativeToStream - nStartRelativeToStream
                                        
                                ' get estimated current message time back to 100% tempo
                                ' already determined
                                dTempo = CDbl(nTempoPrevious) / 600000# * 100# ' percent
                                dTimeDifference100 = dTimeDifferenceOld _
                                 * (1# / (dTempo / 100#))
                                ' 1/x from other to 100%
                                
                                ' get estimated starting time of new stream
                                nTempoCurrent = arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount)
                                dTempo = CDbl(nTempoCurrent) / 600000# * 100# ' percent
                                dTimeDifferenceNew = dTimeDifference100 _
                                 * (1# * (dTempo / 100#))
                                ' 1*x from 100% to other
                                nStartRelativeToStream = nCurrentRelativeToStream - dTimeDifferenceNew
                    
                                nStreamTimeStart = nStreamTimeStart + nStartRelativeToStream
                                nTempoPrevious = nTempoCurrent
                            '}
                            
                            ' New tempo.
                            dTicksPerMillisecond = (CDbl(arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount)) / CDbl(nTempoCurrent)) * 1000#
                            'dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000# ' not applicable
                        
                        ElseIf isMsgFF88TPQChange = True Then
                            ' New time signature.
                            ' Change tick scale but not speed of music.
                            nTempoCurrent = arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount)
                            dTicksPerMillisecond = (CDbl(arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount)) / CDbl(nTempoCurrent)) * 1000#
                            'dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000# ' not applicable
                        End If
                    
                        nTicksRemaining = nTicksRemaining - nTicksBetweenEvents
                    
                        'Exit Do
                    Loop
                    
                    ' Get next message
                    ' store in variables for speed
                    tempmessagestate = MIDIMESSAGESTATE_ENABLED
                    templogonly = False
                    tempmessage = MIDIFile1.Message
                    tempdata1 = MIDIFile1.Data1
                    tempdata2 = MIDIFile1.Data2
                    
                    ' Tag notes to play on keyboard and VU meters
                    tempmessagetag = 0
                    If (tempmessage And &HF0) = note_on And tempdata2 <> 0 Then
                        tempmessagetag = tempdata2 + 1& + (mTrackLogical * 1000&)
                    End If
                    
                    ' Get next time
                    ' Assuming all previous ticks were for one tempo only.
                    ' Assuming shifting start time already compensated for different tempos.
                    nTicksBetweenEvents = nTicksRemaining
                    nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                    nStreamTimeCurrent = RoundVB5(nStreamTicksCurrent / dTicksPerMillisecond, 0) ' ticks to time
                    temptime = nStreamTimeStart + nStreamTimeCurrent
        
                    If temptime > nTrackMessageTimeTo Then
                        ' Message for next increment.
                        nTrackMessageTimeNext = temptime
                        Exit For ' do not process message yet
                    End If
        
                    ' Get buffer (no temporary variable for speed)
                    If tempmessage = SYSEX Then ' SYSEX message
                        MIDIOutput1.buffer = Chr(SYSEX) & MIDIFile1.buffer
                    End If
        
                    ' Queue with MessagePointer
                    MIDIOutput1_MP(MIDIMP_MESSAGESTATE) = tempmessagestate
                    MIDIOutput1_MP(MIDIMP_MESSAGE) = tempmessage
                    MIDIOutput1_MP(MIDIMP_DATA1) = tempdata1
                    MIDIOutput1_MP(MIDIMP_DATA2) = tempdata2
                    MIDIOutput1_MP(MIDIMP_TIME) = temptime
                    MIDIOutput1_MP(MIDIMP_MESSAGETAG) = tempmessagetag
                    MIDIOutput1.MessagePointer(MIDIOutput1_MP(0), UBound(MIDIOutput1_MP)) = 0
        
                    MIDIOutput1.MessageLogOnly = templogonly
                    'MIDIOutput1.Buffer = ... already done

                    ' Alternative (slow in fast loop)
                    'MIDIOutput1.Message = tempmessage
                    'MIDIOutput1.Data1 = tempdata1
                    'MIDIOutput1.Data2 = tempdata2
                    'MIDIOutput1.Time = temptime
                    'MIDIOutput1.MessageTag = tempmessagetag
        
#If 1 = 0 Then ' comment out to enable test
                    If (MIDIOutput1.Message And &HF0) = note_on And MIDIOutput1.Data2 <> 0 Then
                    Debug.Print MIDIOutput1.Message _
                     , MidiNoteString2Display(Chr(MIDIOutput1.Data1)) _
                     , MIDIOutput1.time, Rnd(1)
                    End If
    
                    If nMessageCount = 190 Then Stop ' prevent overflow in debug window
#End If
    
                    ' Add to output queue
                    MIDIOutput1.StreamMessageNumber = 0 ' append
                    MIDIOutput1.ActionStream = MIDIOUT_QUEUE
                    nMessageCount = nMessageCount + 1
                    nTrackMessageCurrent = nTrackMessageCurrent + 1
                    nTrackMessageIncrement = nTrackMessageIncrement + 1

                    ' Update last known processed data.
                    arTrackMessageCurrent(mTrackLogical) = nTrackMessageCurrent
                    arStreamTicksCurrent(mTrackLogical) = nStreamTicksCurrent
                    arStreamTimeCurrent(mTrackLogical) = nStreamTimeCurrent
                    arStreamTimeStart(mTrackLogical) = nStreamTimeStart
                    arTicksPerMillisecond(mTrackLogical) = dTicksPerMillisecond
                    arTempoPrevious(mTrackLogical) = nTempoPrevious
                    arMsgFF81TempoCount(mTrackLogical) = mMsgFF81TempoCount
                    arMsgFF88TPQCount(mTrackLogical) = mMsgFF88TPQCount

#If 1 = 0 Then ' comment out to enable test
                    ' Verify messages are loaded as expected.
                    ' Compare loading by pieces, to by tracks which is more accurate.
                    If mTrackLogical = 2 Then ' any arbitrary track
                    Debug.Print "Track#:"; mTrackLogical _
                    , "To:"; nTrackMessageTimeTo _
                    , "Next:"; nTrackMessageTimeNext _
                    , "Msg#:"; nTrackMessageCurrent _
                    , "Time:"; nStreamTicksCurrent _
                    , nStreamTimeCurrent _
                    , nStreamTimeStart _
                    , temptime _
                    , "Tempo:0"; Left$(Trim$(str$(dTicksPerMillisecond)), 6) _
                    , nTempoPrevious _
                    , mMsgFF81TempoCount _
                    , mMsgFF88TPQCount _
                    , "Max:"; MIDIFile1.MessageCount _
                    , "Msg:"; tempmessage; "-"; tempdata1; "-"; tempdata2
                    'Stop
                    'If nTrackMessageCurrent = 144 Then Stop
                    End If
#End If
            
                    If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' reached limit
                    'If MIDIOutput1.StreamMessageNumber = MB_LONGUBOUND Then Exit For ' not accurate
                    'If nMessageCount = backupstreammessagenumbermax Then Exit For ' not accurate
                Next m
            
                ' Determine next increment group.
                If isTrackMessageEnd = True Then
                    isTrackDone = True
                    'nTrackMessageCurrent = ... ' not needed
                    mTrackLoadComplete = mTrackLoadComplete + 1
                End If
                
                ' Update last known processed data.
                arTrackDone(mTrackLogical) = isTrackDone ' not applicable
                arTrackMessageTimeNext(mTrackLogical) = nTrackMessageTimeNext
            End If ' isTrackMute
    
            If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' reached limit
            'If MIDIOutput1.StreamMessageNumber = MB_LONGUBOUND Then Exit For ' not accurate
            'If nMessageCount = backupstreammessagenumbermax Then Exit For ' not accurate
        Next mTrackPhysical
    
        ' Prepare for next interval
        nTrackMessageTimeCurrent = nTrackMessageTimeCurrent + nTrackMessageTimeIncrement
    
        'Exit Sub
    Loop ' mTrackLoadComplete

#If 1 = 0 Then ' comment out to enable test
    ' Verify number of messages are loaded as expected.
    ' Compare loading by pieces, to by tracks which is more accurate.
    Debug.Print "by pieces", nMessageCount, Rnd()
#End If

    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Sort stream
    ''''''''''''''''''''''''''''''''''''''''''''''
    If MIDIOutput1.StreamMessageSortOutOfOrder = False Then ' already queuesort
    ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STARTED Then ' can only autosort
        Debug.Print "PROGRAM WARNING 21095, autosort"
    Else ' manualsort
        'Debug.Print "PROGRAM WARNING 21095, manualsort"
        'Call MIDIOutput1.SortStreamEx(MIDIOutput1.StreamNumber, 0) ' modal, if stream is small
        Call MIDIOutput1.SortStreamEx(MIDIOutput1.StreamNumber, 1) ' modeless
        Call WaitSortStream(MIDIOutput1.StreamNumber, MIDIOutput1)
    End If
    
    If MIDIOutput1.StreamMessageSortOutOfOrder = True Then ' should have been sorted
    'If MIDIOutput1.StreamMessageSortOutOfOrder = False Then ' want it to be out of order
        Err.Raise 1, , "PROGRAM ERROR 3276"
    End If

    ' Track scrollbar compared to total message time for convenience.
    ' WARNING,
    ' In seconds because data limit is 32767. Milliseconds would
    ' have required a more complex algorithm and data storage.
    HScrollPlayerTime.max = MIDIOutput1.StreamMessageLastTime(1) / 1000
    LabelQueueTime.Caption = Trim$(str$(0))

ExitSection:
    Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
    
    Me.Caption = Me.Tag ' restore
    'Me.Caption = "MFPlayer Example" ' alternative
    Screen.MousePointer = 0 ' backupscreenmousepointer
    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub QueueSong_ByMidi0Tracks()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    '-------------------------------------------------------------------
    ' Queue Song
    ' - scan for global messages, like tempo
    ' - queue one track at a time
    ' - insert and queue global messages, as needed
    ' - queue messages with MessagePointer
    ' - queue message to a stream
    
    ' WARNING,
    ' Queue is slower if sorting a lot of messages
    ' to the middle of a stream. Stays fast if
    ' already sorted near the beginning or end.
    '-------------------------------------------------------------------
    
    Debug.Print "Verifying: Queue midi 0 by tracks, " _
     ; IIf(CheckManualSort.value = 0, "", "sort manually")
    
    Dim backupscreenmousepointer As Integer
    Dim backupstreammessagenumbermax As Long
    Dim mTrackPhysical As Integer
    Dim mTrackLogical As Integer
    Dim isTrackMute As Boolean
    Dim m As Long
    Dim i As Long
    Dim nMessageCount As Long
    Dim nMessageTotal As Long
    Dim mR As Long
    Dim mC As Long
    
    Dim MIDIOutput1_MP(0 To MIDIMP_UBOUND) As Long ' always start from zero
    Dim nMP As Long
    Dim tempmessage As Integer
    Dim tempdata1 As Integer
    Dim tempdata2 As Integer
    Dim temptime As Long
    Dim tempmessagetag As Long
    Dim tempmessagestate As Integer
    Dim templogonly As Boolean
    
    Dim mMsgFF81TempoCount As Integer
    Dim mMsgFF88TPQCount As Integer
    Dim mMsgFF81TempoCountMax As Integer
    Dim mMsgFF88TPQCountMax As Integer
    Dim arMsgFF81Tempo() As Long
    Dim arMsgFF88TPQ() As Long
    Const MB_DIMENSION1UBOUND = 3
    Const MB_TICK = 1
    Const MB_VALUE = 2
    Const MB_TICKNEXT = 3

    Dim backuptempo As Long
    Dim backupticksperquarternote As Integer
    Dim backupnumerator As Integer
    Dim backupdenominator As Integer
    Dim dTicksPerMillisecond As Double
    Dim nTicksBetweenEvents As Long
    Dim nTicksRemaining As Long
    Dim nMillisecondsBetweenEvents As Long
    Dim nStreamTimeCurrent As Long
    Dim nStreamTicksCurrent As Long
    Dim isTrackTicks As Boolean
    Dim nStreamTimeStart As Long
    Dim isGlobal As Boolean
    Dim isMsgFF81TempoChange As Boolean
    Dim isMsgFF88TPQChange As Boolean
    Dim isSortOutOfOrder As Boolean
    
    Dim nStartRelativeToStream As Long
    Dim nCurrentRelativeToStream As Long
    Dim dTimeDifferenceOld As Double
    Dim dTimeDifference100 As Double
    Dim dTimeDifferenceNew As Double
    Dim dTempo As Double
    Dim nTempoCurrent As Long
    Dim nTempoPrevious As Long
    Dim isProcessTempo As Boolean

    If (MIDIFile1.Filename = "") Then GoTo ExitEnd

    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    backupscreenmousepointer = Screen.MousePointer
    Screen.MousePointer = 11
    Me.Tag = Me.Caption ' backup

    ' Pick one stream number the whole form can use.
    Call OpenQueueStream(MainStreamNumber, MB_STREAMNAME_1, MIDIOutput1)
    MIDIOutput1.StreamNumber = MainStreamNumber

    ' Clear any data if stream not new
    MIDIOutput1.ActionStream = MIDIOUT_RESET

    ' Prevent queuesort temporarily for speed
    ' in case too many messages in a midi file
    ' format 1 so not sort on every message.
    ' Assume will sort manually later.
    If CheckManualSort.value = 1 Then MIDIOutput1.StreamMessageSortOutOfOrder = True

    ' Get statistics
    nMessageTotal = 0
    backupstreammessagenumbermax = MIDIOutput1.StreamMessageNumberMax
    For m = 1 To MIDIFile1.NumberOfTracks
        MIDIFile1.TrackNumber = m
        nMessageTotal = nMessageTotal + MIDIFile1.MessageCount
    Next m
    If MIDIFile1.NumberOfTracks = 1 Then
        ' Midi 0 file format
        ' All messages merged and sorted already.
    Else
        ' Midi 1 file format
        ' First track is usually global midi info, but not guarenteed.
        ' Second track is first track with notes.
        ' Etc.
    End If
    
    Me.Caption = "MFPlayer Example - Loading - " _
     & Trim$(str$(Int(100 * nMessageCount / nMessageTotal))) & "%"
    
    ' Get global tags for reference.
    ' Assume in one track if midi file format 0.
    ' Assume in first track if midi file format 1. May be in others but not standard.
    '{
        ' Get tempo info
        MIDIFile1.TrackNumber = 1
        MIDIFile1.MessageNumber = 0
        backuptempo = MIDIFile1.Tempo
        backupticksperquarternote = MIDIFile1.TicksPerQuarterNote
        backupnumerator = MIDIFile1.Numerator
        backupdenominator = 2 ^ MIDIFile1.Denominator
        If backuptempo = 0 Then backuptempo = 600000 ' assume 100 beats per minute (tempo/2 = beats*2)
        If backupticksperquarternote = 0 Then backupticksperquarternote = 480 ' assume 100 beats per minute
        If backupnumerator = 0 Then backupnumerator = 4 ' assume time signature 4/4
        If backupdenominator = 0 Then backupdenominator = 4
        dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000#
        
        ReDim arMsgFF81Tempo(MB_DIMENSION1UBOUND, 0 To 1000)
        ReDim arMsgFF88TPQ(MB_DIMENSION1UBOUND, 0 To 1000)
        
        mMsgFF81TempoCount = 0
        arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) = 0 ' tick zero
        arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount) = backuptempo
        arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) = 0 ' not yet
        
        mMsgFF88TPQCount = 0
        arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) = 0 ' tick zero
        arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount) = backupticksperquarternote
        arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) = 0 ' not yet
        
        nMessageCount = 0
        isSortOutOfOrder = False
        For mTrackPhysical = 1 To MIDIFile1.NumberOfTracks ' 1-based scale
            MIDIFile1.TrackNumber = mTrackPhysical ' 1-based scale (first is global, second is track one)
            mTrackLogical = mTrackPhysical - 1 ' 0-based scale (zero is global, first is track one)
        
            nStreamTicksCurrent = 0
            For m = 1 To MIDIFile1.MessageCount ' 1-based scale
                If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                
                MIDIFile1.MessageNumber = m ' 1-based scale
        
                If Int(m / 1000) = m / 1000 And nMessageTotal <> 0 Then
                    Me.Caption = "MFPlayer Example - Loading - " _
                     & IIf(nMessageCount < nMessageTotal, "0", "") _
                     & Trim$(str$(Int(100 * nMessageCount / nMessageTotal) / 100)) & "%"
                    ' fraction to show preloading
                    Call DoEventsOnce(True): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
                End If
                
                nTicksBetweenEvents = MIDIFile1.time
                nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                
                tempmessage = MIDIFile1.Message
                If tempmessage <> META Then 'ignore
                ElseIf MIDIFile1.Data1 = META_TEMPO Then ' tempo
                    mMsgFF81TempoCount = mMsgFF81TempoCount + 1
                    If mMsgFF81TempoCount > UBound(arMsgFF81Tempo, 2) Then _
                     ReDim Preserve arMsgFF81Tempo(MB_DIMENSION1UBOUND, UBound(arMsgFF81Tempo, 2) + 100) ' more space
                    
                    arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) = nStreamTicksCurrent
                    arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount) = MIDIFile1.Tempo
                    arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) = 0 ' not yet
                    arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount - 1) = nStreamTicksCurrent ' save for reference
                
                    If arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) < arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount - 1) Then _
                     isSortOutOfOrder = True ' when message not in first track
                
                ElseIf MIDIFile1.Data1 = 88 Then ' time sig
                    mMsgFF88TPQCount = mMsgFF88TPQCount + 1
                    If mMsgFF88TPQCount > UBound(arMsgFF88TPQ, 2) Then _
                     ReDim Preserve arMsgFF88TPQ(MB_DIMENSION1UBOUND, UBound(arMsgFF88TPQ, 2) + 100) ' more space
                    
                    arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) = nStreamTicksCurrent
                    arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount) = MIDIFile1.TicksPerQuarterNote
                    arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) = 0 ' not yet
                    arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount - 1) = nStreamTicksCurrent ' save for reference
                
                    If arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) < arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount - 1) Then _
                     isSortOutOfOrder = True ' when message not in first track
                
                End If
            
                nMessageCount = nMessageCount + 1
                
                If nMessageCount = backupstreammessagenumbermax Then Exit For ' reached limit
                'If nMessageCount = MB_LONGUBOUND Then Exit For ' not accurate
                'If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' not applicable
            Next m
            
            If nMessageCount = backupstreammessagenumbermax Then Exit For ' reached limit
            'If nMessageCount = MB_LONGUBOUND Then Exit For ' not accurate
            'If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' not applicable
        Next mTrackPhysical
        
        mMsgFF81TempoCountMax = mMsgFF81TempoCount
        mMsgFF88TPQCountMax = mMsgFF88TPQCount
    
#If 1 = 0 Then ' comment out to enable test
        Debug.Print "------------------------------------------"
        Debug.Print "arMsgFF81Tempo"
        For mR = LBound(arMsgFF81Tempo, 2) To mMsgFF81TempoCountMax
         Debug.Print mR; Space$(5);
         For mC = 1 To UBound(arMsgFF81Tempo, 1)
            Debug.Print arMsgFF81Tempo(mR, mC) & Space$(5);
         Next
         Debug.Print
        Next
        Stop
#End If
    
    '}

    ' Get all messages
    nMessageCount = 0
    For mTrackPhysical = 1 To MIDIFile1.NumberOfTracks ' 1-based scale
        MIDIFile1.TrackNumber = mTrackPhysical ' 1-based scale (first is global, second is track one)
        mTrackLogical = mTrackPhysical - 1 ' 0-based scale (zero is global, first is track one)
        
        isTrackMute = False

#If 1 = 0 Then ' comment out to enable test
        ' Mute some tracks (not channels).
        ' Assuming midi file format 1.
        ' Assuming track zero has all global messages.
        If mTrackLogical <> 0 _
         And mTrackLogical <> 8 _
         And mTrackLogical <> 9 _
         Then
            isTrackMute = True
        End If
#End If
        
        If isTrackMute = False Then
            nStreamTicksCurrent = 0
            nStreamTimeCurrent = 0
            nStreamTimeStart = 0
            nTempoPrevious = backuptempo
            mMsgFF81TempoCount = 0
            mMsgFF88TPQCount = 0
            For m = 1 To MIDIFile1.MessageCount ' 1-based scale
                If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                
                If Int(m / 1000) = m / 1000 And nMessageTotal <> 0 Then
                    Me.Caption = "MFPlayer Example - Loading - " _
                     & Trim$(str$(Int(100 * nMessageCount / nMessageTotal))) & "%"
                    Call DoEventsOnce(True): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
                End If
    
                MIDIFile1.MessageNumber = m ' 1-based scale
            
                ' Get next time
                nTicksRemaining = MIDIFile1.time
                
                ' Insert global messages, if any
                ' if occurs before next message.
                Do
                    ' Assume not scan entire array for tempo since sequential and sorted.
                    isGlobal = False
                    isMsgFF81TempoChange = False
                    isMsgFF88TPQChange = False
                    If mMsgFF81TempoCount <> mMsgFF81TempoCountMax _
                     And nStreamTicksCurrent + nTicksRemaining >= arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) Then
                        isGlobal = True
                        isMsgFF81TempoChange = True
                        mMsgFF81TempoCount = mMsgFF81TempoCount + 1
                        nTicksBetweenEvents = arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) - nStreamTicksCurrent
                    
                    ElseIf mMsgFF88TPQCount <> mMsgFF88TPQCountMax _
                     And nStreamTicksCurrent + nTicksRemaining >= arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) Then
                        isGlobal = True
                        isMsgFF88TPQChange = True
                        mMsgFF88TPQCount = mMsgFF88TPQCount + 1
                        nTicksBetweenEvents = arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) - nStreamTicksCurrent
                    End If
                
                    If isGlobal = False Then Exit Do ' none
                
                    ' Get time
                    ' Assuming all previous ticks were for one tempo only.
                    ' Assuming shifting start time already compensated for.
                    ' Assume tracking ticks is more accurate than tracking time.
                    nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                    nStreamTimeCurrent = RoundVB5(nStreamTicksCurrent / dTicksPerMillisecond, 0) ' ticks to time
                
                    ' Adjust for changes in tempo.
                    If isMsgFF81TempoChange = True Then
                        ' New start time.
                        ' Shift start time to compensate based on
                        ' current time and tempo rate.
                        '{
                            ' estimated message current time
                            nStartRelativeToStream = 0
                            nCurrentRelativeToStream = nStreamTimeCurrent
                            dTimeDifferenceOld = nCurrentRelativeToStream - nStartRelativeToStream
                                    
                            ' get estimated current message time back to 100% tempo
                            ' already determined
                            dTempo = CDbl(nTempoPrevious) / 600000# * 100# ' percent
                            dTimeDifference100 = dTimeDifferenceOld _
                             * (1# / (dTempo / 100#))
                            ' 1/x from other to 100%
                            
                            ' get estimated starting time of new stream
                            nTempoCurrent = arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount)
                            dTempo = CDbl(nTempoCurrent) / 600000# * 100# ' percent
                            dTimeDifferenceNew = dTimeDifference100 _
                             * (1# * (dTempo / 100#))
                            ' 1*x from 100% to other
                            nStartRelativeToStream = nCurrentRelativeToStream - dTimeDifferenceNew
                
                            nStreamTimeStart = nStreamTimeStart + nStartRelativeToStream
                            nTempoPrevious = nTempoCurrent
                        '}
                        
                        ' New tempo.
                        dTicksPerMillisecond = (CDbl(arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount)) / CDbl(nTempoCurrent)) * 1000#
                        'dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000# ' not applicable
                    
                    ElseIf isMsgFF88TPQChange = True Then
                        ' New time signature.
                        ' Change tick scale but not speed of music.
                        nTempoCurrent = arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount)
                        dTicksPerMillisecond = (CDbl(arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount)) / CDbl(nTempoCurrent)) * 1000#
                        'dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000# ' not applicable
                    End If
                
                    nTicksRemaining = nTicksRemaining - nTicksBetweenEvents
                
                    'Exit Do
                Loop
                
                ' Get next message
                ' store in variables for speed
                tempmessagestate = MIDIMESSAGESTATE_ENABLED
                templogonly = False
                tempmessage = MIDIFile1.Message
                tempdata1 = MIDIFile1.Data1
                tempdata2 = MIDIFile1.Data2
                
                ' Tag notes to play on keyboard and VU meters
                tempmessagetag = 0
                If (tempmessage And &HF0) = note_on And tempdata2 <> 0 Then
                    tempmessagetag = tempdata2 + 1& + (mTrackLogical * 1000&)
                End If
                
                ' Get next time
                ' Assuming all previous ticks were for one tempo only.
                ' Assuming shifting start time already compensated for different tempos.
                nTicksBetweenEvents = nTicksRemaining
                nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                nStreamTimeCurrent = RoundVB5(nStreamTicksCurrent / dTicksPerMillisecond, 0) ' ticks to time
                temptime = nStreamTimeStart + nStreamTimeCurrent
    
                ' Get buffer (no temporary variable for speed)
                If tempmessage = SYSEX Then ' SYSEX message
                    MIDIOutput1.buffer = Chr(SYSEX) & MIDIFile1.buffer
                End If
    
                ' Queue with MessagePointer
                MIDIOutput1_MP(MIDIMP_MESSAGESTATE) = tempmessagestate
                MIDIOutput1_MP(MIDIMP_MESSAGE) = tempmessage
                MIDIOutput1_MP(MIDIMP_DATA1) = tempdata1
                MIDIOutput1_MP(MIDIMP_DATA2) = tempdata2
                MIDIOutput1_MP(MIDIMP_TIME) = temptime
                MIDIOutput1_MP(MIDIMP_MESSAGETAG) = tempmessagetag
                MIDIOutput1.MessagePointer(MIDIOutput1_MP(0), UBound(MIDIOutput1_MP)) = 0
    
                MIDIOutput1.MessageLogOnly = templogonly
                'MIDIOutput1.Buffer = ... already done

                ' Alternative (slow in fast loop)
                'MIDIOutput1.Message = tempmessage
                'MIDIOutput1.Data1 = tempdata1
                'MIDIOutput1.Data2 = tempdata2
                'MIDIOutput1.Time = temptime
                'MIDIOutput1.MessageTag = tempmessagetag
    
#If 1 = 0 Then ' comment out to enable test
                If (MIDIOutput1.Message And &HF0) = note_on And MIDIOutput1.Data2 <> 0 Then
                Debug.Print MIDIOutput1.Message _
                 , MidiNoteString2Display(Chr(MIDIOutput1.Data1)) _
                 , MIDIOutput1.time, Rnd(1)
                End If

                If nMessageCount = 190 Then Stop ' prevent overflow in debug window
#End If

                ' Add to output queue
                MIDIOutput1.StreamMessageNumber = 0 ' append
                MIDIOutput1.ActionStream = MIDIOUT_QUEUE
                nMessageCount = nMessageCount + 1

#If 1 = 0 Then ' comment out to enable test
                ' Verify messages are loaded as expected.
                ' Compare loading by pieces, to by tracks which is more accurate.
                If mTrackLogical = 2 Then ' any arbitrary track
                Debug.Print "Track#:"; mTrackLogical _
                , "To:(n/a)" _
                , "Next:(n/a)" _
                , "Msg#:"; m _
                , "Time:"; nStreamTicksCurrent _
                , nStreamTimeCurrent _
                , nStreamTimeStart _
                , temptime _
                , "Tempo:0"; Left$(Trim$(str$(dTicksPerMillisecond)), 6) _
                , nTempoPrevious _
                , mMsgFF81TempoCount _
                , mMsgFF88TPQCount _
                , "Max:"; MIDIFile1.MessageCount _
                , "Msg:"; tempmessage; "-"; tempdata1; "-"; tempdata2
                'Stop
                'If m = 144 Then Stop
                End If
#End If

                If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' reached limit
                'If MIDIOutput1.StreamMessageNumber = MB_LONGUBOUND Then Exit For ' not accurate
                'If nMessageCount = backupstreammessagenumbermax Then Exit For ' not accurate
            Next m
        End If ' isTrackMute
    
        If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' reached limit
        'If MIDIOutput1.StreamMessageNumber = MB_LONGUBOUND Then Exit For ' not accurate
        'If nMessageCount = backupstreammessagenumbermax Then Exit For ' not accurate
    Next mTrackPhysical

#If 1 = 0 Then ' comment out to enable test
    ' Verify number of messages are loaded as expected.
    ' Compare loading by pieces, to by tracks which is more accurate.
    Debug.Print "by track", nMessageCount, Rnd()
#End If

    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Sort stream
    ''''''''''''''''''''''''''''''''''''''''''''''
    If MIDIOutput1.StreamMessageSortOutOfOrder = False Then ' already queuesort
    ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STARTED Then ' can only autosort
        Debug.Print "PROGRAM WARNING 21095, autosort"
    Else ' manualsort
        'Debug.Print "PROGRAM WARNING 21095, manualsort"
        'Call MIDIOutput1.SortStreamEx(MIDIOutput1.StreamNumber, 0) ' modal, if stream is small
        Call MIDIOutput1.SortStreamEx(MIDIOutput1.StreamNumber, 1) ' modeless
        Call WaitSortStream(MIDIOutput1.StreamNumber, MIDIOutput1)
    End If
    
    If MIDIOutput1.StreamMessageSortOutOfOrder = True Then ' should have been sorted
    'If MIDIOutput1.StreamMessageSortOutOfOrder = False Then ' want it to be out of order
        Err.Raise 1, , "PROGRAM ERROR 3276"
    End If

    ' Track scrollbar compared to total message time for convenience.
    ' WARNING,
    ' In seconds because data limit is 32767. Milliseconds would
    ' have required a more complex algorithm and data storage.
    HScrollPlayerTime.max = MIDIOutput1.StreamMessageLastTime(1) / 1000
    LabelQueueTime.Caption = Trim$(str$(0))
    
ExitSection:
    Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
    
    Me.Caption = Me.Tag ' restore
    'Me.Caption = "MFPlayer Example" ' alternative
    Screen.MousePointer = 0 ' backupscreenmousepointer
    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub QueueSong_ByMidi1Track()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    '-------------------------------------------------------------------
    ' Queue Song
    ' - scan for global messages, like tempo
    ' - queue each track into individual streams
    '   to speed up queuesort, manualsort, autosort
    ' - insert and queue global messages, as needed
    ' - queue messages with MessagePointer
    ' - queue message to a stream
    
    ' WARNING,
    ' The total number of tracks will be misleading
    ' if some tracks are empty. The total should be
    ' less. However, for simplicity, they will just be
    ' considered regular tracks with no messages
    ' so that keeping track of them is easier.
    
    ' WARNING,
    ' A midi 0 file format is left in one track.
    ' Converting a midi 0 file format to midi 1
    ' usually should use one track for each channel
    ' and one track for the global channel.
    '-------------------------------------------------------------------
    
    Debug.Print "Verifying: Queue midi 1 by tracks, " _
     ; IIf(CheckManualSort.value = 0, "", "sort manually")
    
    Dim i As Long
    Dim backupscreenmousepointer As Integer
    Dim backupstreammessagenumbermax As Long
    Dim mGroupNumber As Integer
    Dim mStreamNumber As Integer
    Dim nLo As Long
    Dim nHi As Long
    Dim isEmpty As Boolean
    
    Dim mTrackPhysical As Integer
    Dim mTrackLogical As Integer
    Dim isTrackMute As Boolean
    Dim m As Long
    Dim nMessageCount As Long
    Dim nMessageTotal As Long
    Dim mR As Long
    Dim mC As Long
    
    Dim MIDIOutput1_MP(0 To MIDIMP_UBOUND) As Long ' always start from zero
    Dim nMP As Long
    Dim tempmessage As Integer
    Dim tempdata1 As Integer
    Dim tempdata2 As Integer
    Dim temptime As Long
    Dim tempmessagetag As Long
    Dim tempmessagestate As Integer
    Dim templogonly As Boolean
    
    Dim mMsgFF81TempoCount As Integer
    Dim mMsgFF88TPQCount As Integer
    Dim mMsgFF81TempoCountMax As Integer
    Dim mMsgFF88TPQCountMax As Integer
    Dim arMsgFF81Tempo() As Long
    Dim arMsgFF88TPQ() As Long
    Const MB_DIMENSION1UBOUND = 3
    Const MB_TICK = 1
    Const MB_VALUE = 2
    Const MB_TICKNEXT = 3

    Dim backuptempo As Long
    Dim backupticksperquarternote As Integer
    Dim backupnumerator As Integer
    Dim backupdenominator As Integer
    Dim dTicksPerMillisecond As Double
    Dim nTicksBetweenEvents As Long
    Dim nTicksRemaining As Long
    Dim nMillisecondsBetweenEvents As Long
    Dim nStreamTimeCurrent As Long
    Dim nStreamTicksCurrent As Long
    Dim isTrackTicks As Boolean
    Dim nStreamTimeStart As Long
    Dim isGlobal As Boolean
    Dim isMsgFF81TempoChange As Boolean
    Dim isMsgFF88TPQChange As Boolean
    Dim isSortOutOfOrder As Boolean
    
    Dim nStartRelativeToStream As Long
    Dim nCurrentRelativeToStream As Long
    Dim dTimeDifferenceOld As Double
    Dim dTimeDifference100 As Double
    Dim dTimeDifferenceNew As Double
    Dim dTempo As Double
    Dim nTempoCurrent As Long
    Dim nTempoPrevious As Long
    Dim isProcessTempo As Boolean

    If (MIDIFile1.Filename = "") Then GoTo ExitEnd

    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    backupscreenmousepointer = Screen.MousePointer
    Screen.MousePointer = 11
    Me.Tag = Me.Caption ' backup

    ' Pick a group of stream numbers the whole form can use.
    ' (More practical to do this later while collecting data
    ' but would get more complicated.)
    ReDim MainStreamGroup(MIDIFile1.NumberOfTracks + 1, 2) ' 1-based scale, plus master track
    If UBound(MainStreamGroup, 1) = 0 Then Err.Raise 1, , "Missing stream number."
    For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
        isEmpty = False
        If mGroupNumber <= MIDIFile1.NumberOfTracks Then
            MIDIFile1.TrackNumber = mGroupNumber
            If MIDIFile1.MessageCount = 0 Then
                ' empty track
                isEmpty = True
            End If
        Else
            ' master track used as is
        End If
        MainStreamGroup(mGroupNumber, MB_STREAMEMPTY) = isEmpty
        
        'If isEmpty = False Then ' skip if empty, but too complicated to handle later
        mStreamNumber = 0 ' new stream
        Call OpenQueueStream(mStreamNumber, MB_STREAMNAME_1, MIDIOutput1)
        MainStreamGroup(mGroupNumber, MB_STREAMNUMBER) = mStreamNumber
        MIDIOutput1.StreamNumber = mStreamNumber
        
        ' Clear any data if stream not new
        MIDIOutput1.ActionStream = MIDIOUT_RESET
        
        ' Total for reference.
        ' Incl. global track, master track and empty track.
        LabelNumberOfTracks.Caption = Trim$(str$(Val(LabelNumberOfTracks.Caption) + 1))
    
        ' Prevent queuesort temporarily for speed
        ' in case too many messages in a midi file
        ' format 1 so not sort on every message.
        ' Assume will sort manually later.
        If CheckManualSort.value = 1 Then MIDIOutput1.StreamMessageSortOutOfOrder = True
        ' (optional since midi file format 1 already sorted by track)
    Next mGroupNumber
    
    ' Get statistics
    nMessageTotal = 0
    backupstreammessagenumbermax = MIDIOutput1.StreamMessageNumberMax
    For m = 1 To MIDIFile1.NumberOfTracks
        MIDIFile1.TrackNumber = m
        nMessageTotal = nMessageTotal + MIDIFile1.MessageCount
    Next m
    If MIDIFile1.NumberOfTracks = 1 Then
        ' Midi 0 file format
        ' All messages merged and sorted already.
    Else
        ' Midi 1 file format
        ' First track is usually global midi info, but not guarenteed.
        ' Second track is first track with notes.
        ' Etc.
    End If
    
    Me.Caption = "MFPlayer Example - Loading - " _
     & Trim$(str$(Int(100 * nMessageCount / nMessageTotal))) & "%"
    
    ' Get global tags for reference.
    ' Assume in one track if midi file format 0.
    ' Assume in first track if midi file format 1. May be in others but not standard.
    '{
        ' Get tempo info
        MIDIFile1.TrackNumber = 1
        MIDIFile1.MessageNumber = 0
        backuptempo = MIDIFile1.Tempo
        backupticksperquarternote = MIDIFile1.TicksPerQuarterNote
        backupnumerator = MIDIFile1.Numerator
        backupdenominator = 2 ^ MIDIFile1.Denominator
        If backuptempo = 0 Then backuptempo = 600000 ' assume 100 beats per minute (tempo/2 = beats*2)
        If backupticksperquarternote = 0 Then backupticksperquarternote = 480 ' assume 100 beats per minute
        If backupnumerator = 0 Then backupnumerator = 4 ' assume time signature 4/4
        If backupdenominator = 0 Then backupdenominator = 4
        dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000#
        
        ReDim arMsgFF81Tempo(MB_DIMENSION1UBOUND, 0 To 1000)
        ReDim arMsgFF88TPQ(MB_DIMENSION1UBOUND, 0 To 1000)
        
        mMsgFF81TempoCount = 0
        arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) = 0 ' tick zero
        arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount) = backuptempo
        arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) = 0 ' not yet
        
        mMsgFF88TPQCount = 0
        arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) = 0 ' tick zero
        arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount) = backupticksperquarternote
        arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) = 0 ' not yet
        
        nMessageCount = 0
        isSortOutOfOrder = False
        For mTrackPhysical = 1 To MIDIFile1.NumberOfTracks ' 1-based scale
            MIDIFile1.TrackNumber = mTrackPhysical ' 1-based scale (first is global, second is track one)
            mTrackLogical = mTrackPhysical - 1 ' 0-based scale (zero is global, first is track one)
        
            nStreamTicksCurrent = 0
            For m = 1 To MIDIFile1.MessageCount ' 1-based scale
                If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                
                MIDIFile1.MessageNumber = m ' 1-based scale
        
                If Int(m / 1000) = m / 1000 And nMessageTotal <> 0 Then
                    Me.Caption = "MFPlayer Example - Loading - " _
                     & IIf(nMessageCount < nMessageTotal, "0", "") _
                     & Trim$(str$(Int(100 * nMessageCount / nMessageTotal) / 100)) & "%"
                    ' fraction to show preloading
                    Call DoEventsOnce(True): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
                End If
                
                nTicksBetweenEvents = MIDIFile1.time
                nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                
                tempmessage = MIDIFile1.Message
                tempdata1 = MIDIFile1.Data1
                If tempmessage <> META Then 'ignore
                ElseIf tempdata1 = META_TEMPO Then ' tempo
                    mMsgFF81TempoCount = mMsgFF81TempoCount + 1
                    If mMsgFF81TempoCount > UBound(arMsgFF81Tempo, 2) Then _
                     ReDim Preserve arMsgFF81Tempo(MB_DIMENSION1UBOUND, UBound(arMsgFF81Tempo, 2) + 100) ' more space
                    
                    arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) = nStreamTicksCurrent
                    arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount) = MIDIFile1.Tempo
                    arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) = 0 ' not yet
                    arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount - 1) = nStreamTicksCurrent ' save for reference
                
                    If arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) < arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount - 1) Then _
                     isSortOutOfOrder = True ' when message not in first track
                
                ElseIf tempdata1 = 88 Then ' time sig
                    mMsgFF88TPQCount = mMsgFF88TPQCount + 1
                    If mMsgFF88TPQCount > UBound(arMsgFF88TPQ, 2) Then _
                     ReDim Preserve arMsgFF88TPQ(MB_DIMENSION1UBOUND, UBound(arMsgFF88TPQ, 2) + 100) ' more space
                    
                    arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) = nStreamTicksCurrent
                    arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount) = MIDIFile1.TicksPerQuarterNote
                    arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) = 0 ' not yet
                    arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount - 1) = nStreamTicksCurrent ' save for reference
                
                    If arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) < arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount - 1) Then _
                     isSortOutOfOrder = True ' when message not in first track
                
                End If
            
                nMessageCount = nMessageCount + 1
                
                If nMessageCount = backupstreammessagenumbermax Then Exit For ' reached limit
                'If nMessageCount = MB_LONGUBOUND Then Exit For ' not accurate
                'If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' not applicable
            Next m
            
            If nMessageCount = backupstreammessagenumbermax Then Exit For ' reached limit
            'If nMessageCount = MB_LONGUBOUND Then Exit For ' not accurate
            'If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' not applicable
        Next mTrackPhysical
        
        mMsgFF81TempoCountMax = mMsgFF81TempoCount
        mMsgFF88TPQCountMax = mMsgFF88TPQCount

    
#If 1 = 0 Then ' comment out to enable test
        Debug.Print "------------------------------------------"
        Debug.Print "arMsgFF81Tempo"
        For mR = LBound(arMsgFF81Tempo, 2) To mMsgFF81TempoCountMax
         Debug.Print mR; Space$(5);
         For mC = 1 To UBound(arMsgFF81Tempo, 1)
            Debug.Print arMsgFF81Tempo(mR, mC) & Space$(5);
         Next
         Debug.Print
        Next
        Stop
#End If
    
    '}
    
    ' Get all messages
    nMessageCount = 0
    For mTrackPhysical = 1 To MIDIFile1.NumberOfTracks ' 1-based scale
        MIDIFile1.TrackNumber = mTrackPhysical ' 1-based scale (first is global, second is track one)
        mTrackLogical = mTrackPhysical - 1 ' 0-based scale (zero is global, first is track one)
        
        mGroupNumber = mTrackPhysical ' 1-based scale (zero is none, first is global, second is track one, last is master track)
        'For mGroupNumber = 1 To UBound(MainStreamGroup, 1) ' not applicable
        MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        
        isTrackMute = False

#If 1 = 0 Then ' comment out to enable test
        ' Mute some tracks (not channels).
        ' Assuming midi file format 1.
        ' Assuming track zero has all global messages.
        If mTrackLogical <> 0 _
         And mTrackLogical <> 8 _
         And mTrackLogical <> 9 _
         Then
            isTrackMute = True
        End If
#End If
        
        If isTrackMute = False Then
            nStreamTicksCurrent = 0
            nStreamTimeCurrent = 0
            nStreamTimeStart = 0
            nTempoPrevious = backuptempo
            mMsgFF81TempoCount = 0
            mMsgFF88TPQCount = 0
            For m = 1 To MIDIFile1.MessageCount ' 1-based scale
                If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                
                If Int(m / 1000) = m / 1000 And nMessageTotal <> 0 Then
                    Me.Caption = "MFPlayer Example - Loading - " _
                     & Trim$(str$(Int(100 * nMessageCount / nMessageTotal))) & "%"
                    Call DoEventsOnce(True): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
                End If
    
                MIDIFile1.MessageNumber = m ' 1-based scale
            
                ' Get next time
                nTicksRemaining = MIDIFile1.time
                
                ' Insert global messages, if any
                ' if occurs before next message.
                Do
                    ' Assume not scan entire array for tempo since sequential and sorted.
                    isGlobal = False
                    isMsgFF81TempoChange = False
                    isMsgFF88TPQChange = False
                    If mMsgFF81TempoCount <> mMsgFF81TempoCountMax _
                     And nStreamTicksCurrent + nTicksRemaining >= arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) Then
                        isGlobal = True
                        isMsgFF81TempoChange = True
                        mMsgFF81TempoCount = mMsgFF81TempoCount + 1
                        nTicksBetweenEvents = arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) - nStreamTicksCurrent
                    
                    ElseIf mMsgFF88TPQCount <> mMsgFF88TPQCountMax _
                     And nStreamTicksCurrent + nTicksRemaining >= arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) Then
                        isGlobal = True
                        isMsgFF88TPQChange = True
                        mMsgFF88TPQCount = mMsgFF88TPQCount + 1
                        nTicksBetweenEvents = arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) - nStreamTicksCurrent
                    End If
                
                    If isGlobal = False Then Exit Do ' none
                
                    ' Get time
                    ' Assuming all previous ticks were for one tempo only.
                    ' Assuming shifting start time already compensated for.
                    ' Assume tracking ticks is more accurate than tracking time.
                    nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                    nStreamTimeCurrent = RoundVB5(nStreamTicksCurrent / dTicksPerMillisecond, 0) ' ticks to time
                
                    ' Adjust for changes in tempo.
                    If isMsgFF81TempoChange = True Then
                        ' New start time.
                        ' Shift start time to compensate based on
                        ' current time and tempo rate.
                        '{
                            ' estimated message current time
                            nStartRelativeToStream = 0
                            nCurrentRelativeToStream = nStreamTimeCurrent
                            dTimeDifferenceOld = nCurrentRelativeToStream - nStartRelativeToStream
                                    
                            ' get estimated current message time back to 100% tempo
                            ' already determined
                            dTempo = CDbl(nTempoPrevious) / 600000# * 100# ' percent
                            dTimeDifference100 = dTimeDifferenceOld _
                             * (1# / (dTempo / 100#))
                            ' 1/x from other to 100%
                            
                            ' get estimated starting time of new stream
                            nTempoCurrent = arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount)
                            dTempo = CDbl(nTempoCurrent) / 600000# * 100# ' percent
                            dTimeDifferenceNew = dTimeDifference100 _
                             * (1# * (dTempo / 100#))
                            ' 1*x from 100% to other
                            nStartRelativeToStream = nCurrentRelativeToStream - dTimeDifferenceNew
                
                            nStreamTimeStart = nStreamTimeStart + nStartRelativeToStream
                            nTempoPrevious = nTempoCurrent
                        '}
                        
                        ' New tempo.
                        dTicksPerMillisecond = (CDbl(arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount)) / CDbl(nTempoCurrent)) * 1000#
                        'dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000# ' not applicable
                    
                    ElseIf isMsgFF88TPQChange = True Then
                        ' New time signature.
                        ' Change tick scale but not speed of music.
                        nTempoCurrent = arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount)
                        dTicksPerMillisecond = (CDbl(arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount)) / CDbl(nTempoCurrent)) * 1000#
                        'dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000# ' not applicable
                    End If
                
                    nTicksRemaining = nTicksRemaining - nTicksBetweenEvents
                
                    'Exit Do
                Loop
                
                ' Get next message
                ' store in variables for speed
                tempmessagestate = MIDIMESSAGESTATE_ENABLED
                templogonly = False
                tempmessage = MIDIFile1.Message
                tempdata1 = MIDIFile1.Data1
                tempdata2 = MIDIFile1.Data2
                
                ' Tag notes to play on keyboard and VU meters
                tempmessagetag = 0
                If (tempmessage And &HF0) = note_on And tempdata2 <> 0 Then
                    tempmessagetag = tempdata2 + 1& + (mTrackLogical * 1000&)
                End If
                
                ' Get next time
                ' Assuming all previous ticks were for one tempo only.
                ' Assuming shifting start time already compensated for different tempos.
                nTicksBetweenEvents = nTicksRemaining
                nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                nStreamTimeCurrent = RoundVB5(nStreamTicksCurrent / dTicksPerMillisecond, 0) ' ticks to time
                temptime = nStreamTimeStart + nStreamTimeCurrent
    
                ' Get buffer (no temporary variable for speed)
                If tempmessage = SYSEX Then ' SYSEX message
                    MIDIOutput1.buffer = Chr(SYSEX) & MIDIFile1.buffer
                End If
    
                ' Queue with MessagePointer
                MIDIOutput1_MP(MIDIMP_MESSAGESTATE) = tempmessagestate
                MIDIOutput1_MP(MIDIMP_MESSAGE) = tempmessage
                MIDIOutput1_MP(MIDIMP_DATA1) = tempdata1
                MIDIOutput1_MP(MIDIMP_DATA2) = tempdata2
                MIDIOutput1_MP(MIDIMP_TIME) = temptime
                MIDIOutput1_MP(MIDIMP_MESSAGETAG) = tempmessagetag
                MIDIOutput1.MessagePointer(MIDIOutput1_MP(0), UBound(MIDIOutput1_MP)) = 0
    
                MIDIOutput1.MessageLogOnly = templogonly
                'MIDIOutput1.Buffer = ... already done

                ' Alternative (slow in fast loop)
                'MIDIOutput1.MessageState = tempmessage
                'MIDIOutput1.Message = tempmessage
                'MIDIOutput1.Data1 = tempdata1
                'MIDIOutput1.Data2 = tempdata2
                'MIDIOutput1.Time = temptime
                'MIDIOutput1.MessageTag = tempmessagetag
    
#If 1 = 0 Then ' comment out to enable test
                If (MIDIOutput1.Message And &HF0) = note_on And MIDIOutput1.Data2 <> 0 Then
                Debug.Print MIDIOutput1.Message _
                 , MidiNoteString2Display(Chr(MIDIOutput1.Data1)) _
                 , MIDIOutput1.time, Rnd(1)
                End If

                If nMessageCount = 190 Then Stop ' prevent overflow in debug window
#End If

                ' Add to output queue
                MIDIOutput1.StreamMessageNumber = 0 ' append
                MIDIOutput1.ActionStream = MIDIOUT_QUEUE
                nMessageCount = nMessageCount + 1

#If 1 = 0 Then ' comment out to enable test
                ' Verify messages are loaded as expected.
                ' Compare loading by pieces, to by tracks which is more accurate.
                If mTrackLogical = 2 Then ' any arbitrary track
                Debug.Print "Track#:"; mTrackLogical _
                , "To:(n/a)" _
                , "Next:(n/a)" _
                , "Msg#:"; m _
                , "Time:"; nStreamTicksCurrent _
                , nStreamTimeCurrent _
                , nStreamTimeStart _
                , temptime _
                , "Tempo:0"; Left$(Trim$(str$(dTicksPerMillisecond)), 6) _
                , nTempoPrevious _
                , mMsgFF81TempoCount _
                , mMsgFF88TPQCount _
                , "Max:"; MIDIFile1.MessageCount _
                , "Msg:"; tempmessage; "-"; tempdata1; "-"; tempdata2
                'Stop
                'If m = 144 Then Stop
                End If
#End If

                If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' reached limit
                'If MIDIOutput1.StreamMessageNumber = MB_LONGUBOUND Then Exit For ' not accurate
                'If nMessageCount = backupstreammessagenumbermax Then Exit For ' not accurate
            Next m
        End If ' isTrackMute
    
        If MIDIOutput1.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' reached limit
        'If MIDIOutput1.StreamMessageNumber = MB_LONGUBOUND Then Exit For ' not accurate
        'If nMessageCount = backupstreammessagenumbermax Then Exit For ' not accurate
    Next mTrackPhysical

#If 1 = 0 Then ' comment out to enable test
    ' Verify number of messages are loaded as expected.
    ' Compare loading by pieces, to by tracks which is more accurate.
    Debug.Print "by track", nMessageCount, Rnd()
#End If

    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Build master track with any extra data to help
    ' in case midi file did not have a global track.
    ' Usually for global channels (other channels optional)
    ''''''''''''''''''''''''''''''''''''''''''''''
    mGroupNumber = UBound(MainStreamGroup, 1) ' last is master track
    MainStreamGroup(mGroupNumber, MB_STREAMNUMBER) = mGroupNumber
    'For mGroupNumber = 1 To UBound(MainStreamGroup, 1) ' not applicable
    MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
    
    ' First message to describe stream for reference (optional)
    MIDIOutput1.MessageState = MIDIMESSAGESTATE_ENABLED
    MIDIOutput1.MessageLogOnly = True
    MIDIOutput1.Message = META
    MIDIOutput1.Data1 = META_MARKER ' pass type of marker (0 to 255)
    MIDIOutput1.Data2 = 0 ' pass information (optional)
    MIDIOutput1.buffer = "built at, " & time
    MIDIOutput1.time = 0
    MIDIOutput1.MessageTag = 0
    MIDIOutput1.StreamMessageNumber = 0 ' append
    MIDIOutput1.ActionStream = MIDIOUT_QUEUE

    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Add common maximum message time for scrollbar or replay
    ''''''''''''''''''''''''''''''''''''''''''''''
    nLo = 0
    nHi = 0
    For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
        MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        
        ' Find common range in time.
        If MIDIOutput1.StreamMessageLastTime(1) > nHi Then _
         nHi = MIDIOutput1.StreamMessageLastTime(1)
    Next mGroupNumber
    
    For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
        MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        
        MIDIOutput1.MessageState = MIDIMESSAGESTATE_ENABLED
        MIDIOutput1.MessageLogOnly = True
        MIDIOutput1.Message = META
        MIDIOutput1.Data1 = META_MARKER ' pass type of marker (0 to 255)
        MIDIOutput1.Data2 = 0 ' pass information (optional)
        MIDIOutput1.buffer = "Lowest time"
        MIDIOutput1.time = nLo
        MIDIOutput1.MessageTag = 0
        MIDIOutput1.StreamMessageNumber = 0 ' append
        MIDIOutput1.ActionStream = MIDIOUT_QUEUE
        
        MIDIOutput1.buffer = "Highest time"
        MIDIOutput1.time = nHi
        MIDIOutput1.StreamMessageNumber = 0 ' append
        MIDIOutput1.ActionStream = MIDIOUT_QUEUE
    Next mGroupNumber

    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Sort stream
    ''''''''''''''''''''''''''''''''''''''''''''''
    For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
        MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        
        If MIDIOutput1.StreamMessageSortOutOfOrder = False Then ' already queuesort
        ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STARTED Then ' can only autosort
            Debug.Print "PROGRAM WARNING 21095, autosort"
        Else ' manualsort
            'Debug.Print "PROGRAM WARNING 21095, manualsort"
            'Call MIDIOutput1.SortStreamEx(MIDIOutput1.StreamNumber, 0) ' modal, if stream is small
            Call MIDIOutput1.SortStreamEx(MIDIOutput1.StreamNumber, 1) ' modeless
            Call WaitSortStream(MIDIOutput1.StreamNumber, MIDIOutput1)
        End If
        
        If MIDIOutput1.StreamMessageSortOutOfOrder = True Then ' should have been sorted
        'If MIDIOutput1.StreamMessageSortOutOfOrder = False Then ' want it to be out of order
            Err.Raise 1, , "PROGRAM ERROR 3276"
        End If
    Next mGroupNumber

    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Track scrollbar compared to total message time for convenience.
    ''''''''''''''''''''''''''''''''''''''''''''''
    HScrollPlayerTime.max = nHi / 1000
    LabelQueueTime.Caption = Trim$(str$(0))
    ' WARNING,
    ' In seconds because data limit is 32767. Milliseconds would
    ' have required a more complex algorithm and data storage.

ExitSection:
    Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
    
    Me.Caption = Me.Tag ' restore
    'Me.Caption = "MFPlayer Example" ' alternative
    Screen.MousePointer = 0 ' backupscreenmousepointer
    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Public Sub StartPlay()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close
    
    Dim mGroupNumber As Integer
    Dim nTime As Long
    
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    
    If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
        If MainStreamNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamNumber
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(0) <> MIDISTATE_STARTED Then
            'ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_OPENED Or MIDIOutput1.StateStreamEx(0) = MIDISTATE_STOPPED Then ' alternative, for backward compatibility with v.1.10.007
                
                'If MIDIOutput1.StreamTimeCurrent = 0 Then ' start from stop
                'If MIDIOutput1.StreamTimeCurrent > 0 Then ' start from pause
                
                ' Start from stop or pause.
                Call CheckAutoReplay_Click
                Call CheckAutoStop_Click
                MIDIOutput1.FilterLateEventStreamMax = True ' may filter notes
                Call CheckMidiOutFilterLateEventAllMax_Click ' may filter notes
                MIDIOutput1.StreamTimeStartRelativeToOpen = MIDIOutput1.TimeRelativeToOpen _
                 - HScrollPlayerTime.value * MB_HSCROLLTIMESCALEOFFSET _
                 * (MIDIOutput1.StreamTempoRate / 100)
                'MIDIOutput1.StreamTimeStartRelativeToOpen = MIDIOutput1.TimeRelativeToOpen _
                ' - ConvertMessageToTime(HScrollPlayerMessage.Value * MB_HSCROLLMESSAGESCALEOFFSET) _
                ' * (MIDIOutput1.StreamTempoRate / 100) ' alternative, but requires conversion
                'MIDIOutput1.StreamTimeStartPending = True ' not needed, use current position
                MIDIOutput1.ActionStream = MIDIOUT_START
            
            ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STARTED Then
                ' Pause with same start button
                ' (not practical)
            End If
        End If

    Else
        ' Midi format 1
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(0) <> MIDISTATE_STARTED Then
            'ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_OPENED Or MIDIOutput1.StateStreamEx(0) = MIDISTATE_STOPPED Then ' alternative, for backward compatibility with v.1.10.007
                
                'If MIDIOutput1.StreamTimeCurrent = 0 Then ' start from stop
                'If MIDIOutput1.StreamTimeCurrent > 0 Then ' start from pause
                
                ' Start from stop or pause.
                Call CheckAutoReplay_Click
                Call CheckAutoStop_Click
                MIDIOutput1.FilterLateEventStreamMax = True ' may filter notes
                Call CheckMidiOutFilterLateEventAllMax_Click ' may filter notes
                If nTime = 0 Then ' get once, same for all streams
                    nTime = MIDIOutput1.TimeRelativeToOpen _
                     - HScrollPlayerTime.value * MB_HSCROLLTIMESCALEOFFSET _
                     * (MIDIOutput1.StreamTempoRate / 100)
                    'MIDIOutput1.StreamTimeStartRelativeToOpen = MIDIOutput1.TimeRelativeToOpen _
                    ' - ConvertMessageToTime(HScrollPlayerMessage.Value * MB_HSCROLLMESSAGESCALEOFFSET) _
                    ' * (MIDIOutput1.StreamTempoRate / 100) ' alternative, but requires conversion
                End If
                MIDIOutput1.StreamTimeStartRelativeToOpen = nTime
                'MIDIOutput1.StreamTimeStartPending = True ' not needed, use current position
                MIDIOutput1.ActionStream = MIDIOUT_START
            
            ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STARTED Then
                ' Pause with same start button
                ' (not practical)
            End If
        Next mGroupNumber
    End If
    
ExitSection:

    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
   
Public Sub PausePlay()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close
    
    Dim mGroupNumber As Integer
    Dim isPause As Boolean
    Dim nTime As Long
    Dim backupmessageventpause As Boolean
    
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    
    If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
        ' Midi format 0
        If MainStreamNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamNumber
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STARTED Then ' started
                ' Pause from start
                MIDIOutput1.ActionStream = MIDIOUT_PAUSE
                'MIDIOutput1.MessageEventPause = True ' alternative, but conditions are different
                Call StopStuckNote
            
            ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STOPPED And MIDIOutput1.StreamTimeCurrent = 0 _
             And CheckPauseRestart = 1 Then
                ' Start from stop with same pause button
                ' (not practical)
            
            ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STOPPED And MIDIOutput1.StreamTimeCurrent > 0 _
             And CheckPauseRestart = 1 Then
                ' Start from stop with same pause button
                ' (optional)
                Call CheckAutoReplay_Click
                Call CheckAutoStop_Click
                MIDIOutput1.FilterLateEventStreamMax = True ' may filter notes
                Call CheckMidiOutFilterLateEventAllMax_Click ' may filter notes
                'MIDIOutput1.StreamTimeStartRelativeToOpen = MIDIOutput1.TimeRelativeToOpen _
                ' - HScrollPlayerTime.Value * MB_HSCROLLTIMESCALEOFFSET _
                ' * (MIDIOutput1.StreamTempoRate / 100)
                'MIDIOutput1.StreamTimeStartRelativeToOpen = MIDIOutput1.TimeRelativeToOpen _
                ' - ConvertMessageToTime(HScrollPlayerMessage.Value * MB_HSCROLLMESSAGESCALEOFFSET) _
                ' * (MIDIOutput1.StreamTempoRate / 100) ' alternative, but requires conversion
                'MIDIOutput1.StreamTimeStartPending = True ' not needed, use current position
                ' (no change in start time because scrollbar is inaccurate)
                MIDIOutput1.ActionStream = MIDIOUT_START
            End If
        End If

    Else
        ' Midi format 1
        isPause = False
        
        ' Pause all streams at exact same time.
        backupmessageventpause = MIDIOutput1.MessageEventPause
        If MIDIOutput1.State <> MIDISTATE_CLOSED Then _
         MIDIOutput1.MessageEventPause = True
        
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STARTED Then ' started
                ' Pause from start
                MIDIOutput1.ActionStream = MIDIOUT_PAUSE
                'MIDIOutput1.MessageEventPause = True ' alternative, but conditions are different
                Call StopStuckNote
                isPause = True
            
            ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STOPPED And MIDIOutput1.StreamTimeCurrent = 0 _
             And CheckPauseRestart = 1 Then
                ' Start from stop with same pause button
                ' (not practical)
            
            ElseIf MIDIOutput1.StateStreamEx(0) = MIDISTATE_STOPPED And MIDIOutput1.StreamTimeCurrent > 0 _
             And CheckPauseRestart = 1 Then
                ' Start from stop with same pause button
                ' (optional)
                Call CheckAutoReplay_Click
                Call CheckAutoStop_Click
                MIDIOutput1.FilterLateEventStreamMax = True ' may filter notes
                Call CheckMidiOutFilterLateEventAllMax_Click ' may filter notes
                If nTime = 0 Then ' get once, same for all streams
                    'MIDIOutput1.StreamTimeStartRelativeToOpen = MIDIOutput1.TimeRelativeToOpen _
                    ' - HScrollPlayerTime.Value * MB_HSCROLLTIMESCALEOFFSET _
                    ' * (MIDIOutput1.StreamTempoRate / 100)
                    'MIDIOutput1.StreamTimeStartRelativeToOpen = MIDIOutput1.TimeRelativeToOpen _
                    ' - ConvertMessageToTime(HScrollPlayerMessage.Value * MB_HSCROLLMESSAGESCALEOFFSET) _
                    ' * (MIDIOutput1.StreamTempoRate / 100) ' alternative, but requires conversion
                    ' (no change in start time because scrollbar is inaccurate)
                End If
                'MIDIOutput1.StreamTimeStartPending = True ' not needed, use current position
                MIDIOutput1.ActionStream = MIDIOUT_START
            End If
        Next mGroupNumber
        If isPause = True Then Call StopStuckNote ' do again in case some streams were still processing
        If MIDIOutput1.State <> MIDISTATE_CLOSED Then _
         MIDIOutput1.MessageEventPause = backupmessageventpause ' restore
    End If

ExitSection:

    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
   
Public Sub StopPlay()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close
    
    Dim mGroupNumber As Integer
    Dim isStop As Boolean
    
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    
    If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
        If MainStreamNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamNumber
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            Else
                MIDIOutput1.ActionStream = MIDIOUT_STOP
                Call StopStuckNote
                Call ClearScrollBar
            End If
        End If

    Else
        ' Midi format 1
        isStop = False
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            Else
                MIDIOutput1.ActionStream = MIDIOUT_STOP
                Call StopStuckNote
                'Call ClearScrollBar ' do after all streams
                isStop = True
            End If
        Next mGroupNumber
        If isStop = True Then
            Call StopStuckNote ' do again in case some streams were still processing
            Call ClearScrollBar
        End If
    End If

ExitSection:

    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
    
Private Sub StopStuckNote()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close
    
    ' Turn all note off
    ' Clear with less obvious option to minimize output.
    'MIDIOutput1.SendNoteOff (1) ' clear all notes off for stuck notes and sustain
    'MIDIOutput1.SendNoteOff (2) ' clear all sounds off for stuck notes and sustain
    'MIDIOutput1.SendNoteOff (3) ' clear each note for stuck notes and sustain
    MIDIOutput1.SendNoteOff (4) ' clear recent notes for stuck notes and sustain

    ' Alternative, turn each note off
    'For y = 0 To 15
    '    For x = 0 To 127
    '        MIDIOutput1.Message = NOTE_ON + y
    '        MIDIOutput1.Data1 = x
    '        MIDIOutput1.Data2 = 0
    '        MIDIOutput1.Action = MIDIOUT_SEND
    '    Next x
    'Next y
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub ShowMidiError( _
   ByVal ErrorCode As Integer _
 , ByVal ErrorMessage As String _
 , ByVal ErrorMessageSource As String _
 , ByVal ErrorCount As String _
 , ByVal isStop As Boolean _
 )
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    '-------------------------------------------------------------------
    ' Normally get the combined message
    ' of the description, source and count, for the most clarification.
    ' Syntax,
    '      Code, Message [Source] (1 of Count errors)
    '      Code, Message [mainprocedure.subprocedure(linenumber)] (1 of x errors)
    '-------------------------------------------------------------------
    
    Dim ErrDescriptionFull As String
    
    ' Show error occurred somewhere before
    ' and was not handled.
    ErrDescriptionFull = ErrorMessage _
     & " [" & ErrorMessageSource & "]"
    If ErrorCount > 1 Then
        ErrDescriptionFull = ErrDescriptionFull _
         & " (1 of " & Trim$(str$(ErrorCount)) & " errors)"
    End If
    
    If isStop = False Then
        Call MsgBoxBug(True) ' prevent multithreading issues caused by msgbox
        MsgBox Trim$(str$(ErrorCode)) & ", " & ErrDescriptionFull
        Call MsgBoxBug(False) ' restore
    Else
        Call MsgBoxBug(True) ' prevent multithreading issues caused by msgbox
        MsgBox Trim$(str$(ErrorCode)) & ", " & ErrDescriptionFull
        Call MsgBoxBug(False) ' restore
        Unload Me
        'On Error GoTo 0 ' does not help disable previous onerror in callstack
        'Err.Raise ErrorCode, , ErrDescriptionFull ' not applicable if previous onerror
    End If
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
   
Private Sub OpenQueueStream( _
   ByRef mStreamNumber As Integer _
 , ByRef cStreamName As String _
 , ByRef MIDIOutput1 As MIDIOutput)
    
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    '-------------------------------------------------------------------
    ' Open queue stream once if not already open by calling program.
    ' If streamnumber is zero, open and return a new streamnumber.
    ' If not zero, then erase streamname.
    '-------------------------------------------------------------------
        
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative

    If Len(Trim$(cStreamName)) = 0 Then _
     Err.Raise 1, , "missing name"
        
    If mStreamNumber = 0 Then
        MIDIOutput1.StreamNumber = mStreamNumber
        MIDIOutput1.StreamName = cStreamName
        If MIDIOutput1.StreamNumberTotal = MIDIOutput1.StreamNumberMax Then _
         Err.Raise 1, , "too many streams"
        MIDIOutput1.ActionStream = MIDIOUT_OPEN
        ' open is the same state as if stopped and on first message number
        mStreamNumber = MIDIOutput1.StreamNumber
    Else
        ' Do not erase since still used for reference later.
        ' But remember it does not match existing stream
        ' and its not supposed to start or close stream later.
        'cStreamName = ""
    End If
    MIDIOutput1.StreamNumber = mStreamNumber
    If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then _
     Err.Raise 1, , "" ' old stream is specified but not open properly ?
    If MIDIOutput1.StreamName = "" Then _
     Err.Raise 1, , ""

    ' mStreamNumber returned ByRef
    ' cStreamName returned ByRef
    
ExitSection:

    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Public Sub WaitStopStream( _
   ByVal mStreamNumber As Integer _
 , ByVal cStreamName As String _
 , ByRef MIDIOutput1 As MIDIOutput)
    
    If gisEnd = True Then GoTo ExitEnd

    '-------------------------------------------------------------------
    ' Wait until all related streams be empty and autostop
    ' if object is enabled. Routine is polling for changes
    ' which is not efficient but sleep helps minimize cpu usage.
    ' The form and application will be accessible during loop so any objects
    ' should have already been disabled that are not supposed to run meanwhile.
    '-------------------------------------------------------------------
        
    Dim isStreamWaitExit As Boolean
    
    ' Preserve passed data so not interfere with other functions
    'Dim backupelement As Integer
    'Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    ' (only use temporarily because not multhread-safe to use while in loop)
    
    If mStreamNumber = 0 And cStreamName = "" Then
        Err.Raise 1, , "PROGRAM ERROR 26115, missing condition"
    End If
    
    isStreamWaitExit = False
    Do While isStreamWaitExit = False
        If gisEnd = True Then GoTo ExitEnd
        
        Sleep MB_DOEVENTSPOLLING ' release resources enough so <5% cpu usage
        Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
        'If MIDIOutput1.StreamNumber <> mStreamNumber Then Call Error3886 ' use if doevents screws it up
        'MIDIOutput1.MessageEventPending = False ' alternative
    
        Do
            If gisEnd = True Then GoTo ExitEnd

            If mStreamNumber = 0 Then
            ElseIf MIDIOutput1.StateStreamEx(mStreamNumber) = MIDISTATE_CLOSED Then
            ElseIf MIDIOutput1.StateStreamEx(mStreamNumber) = MIDISTATE_STOPPED Then
            Else
                Exit Do ' keep checking
            End If

            If cStreamName = "" Then
            ElseIf MIDIOutput1.StateStreamNameEx(cStreamName, mStreamNumber) = MIDISTATE_CLOSED Then ' find nearest stream
            ElseIf MIDIOutput1.StateStreamNameEx(cStreamName, mStreamNumber) = MIDISTATE_STOPPED Then ' find nearest stream
            'ElseIf MIDIOutput1.StateStreamNameEx(MIDISTATE_STARTED, mStreamNumber) = MIDISTATE_CLOSED Then ' find first started stream
            'ElseIf MIDIOutput1.StateStreamNameEx("", mStreamNumber) = MIDISTATE_CLOSED Then ' find first stream not closed
            Else
                Exit Do ' keep checking
            End If
            'MIDIOutput1.StreamNumber = mStreamNumber
            
            ' Wait for other streams too (optional).
            ' Wait based on other criteria (optional).
            
            isStreamWaitExit = True ' no more found
            Exit Do ' only run once
        Loop
        If isStreamWaitExit = True Then
            Exit Do  ' no more needed to check
        End If
    Loop
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Public Sub WaitCloseStream( _
   ByVal mStreamNumber As Integer _
 , ByVal cStreamName As String _
 , ByRef MIDIOutput1 As MIDIOutput)
    
    If gisEnd = True Then GoTo ExitEnd

    '-------------------------------------------------------------------
    ' Wait until all related streams be empty and autoclose
    ' if object is enabled. Routine is polling for changes
    ' which is not efficient but sleep helps minimize cpu usage.
    ' The form and application will be accessible during loop so any objects
    ' should have already been disabled that are not supposed to run meanwhile.
    '-------------------------------------------------------------------
        
    Dim isStreamWaitExit As Boolean
    
    ' Preserve passed data so not interfere with other functions
    'Dim backupelement As Integer
    'Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    ' (only use temporarily because not multhread-safe to use while in loop)
    
    If mStreamNumber = 0 And cStreamName = "" Then
        Err.Raise 1, , "PROGRAM ERROR 26115, missing condition"
    End If
    
    isStreamWaitExit = False
    Do While isStreamWaitExit = False
        If gisEnd = True Then GoTo ExitEnd
        
        Sleep MB_DOEVENTSPOLLING ' release resources enough so <5% cpu usage
        Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
        'MIDIOutput1.MessageEventPending = False ' alternative
        'If MIDIOutput1.StreamNumber <> mStreamNumber Then Call Error3886 ' use if doevents screws it up
    
        Do
            If gisEnd = True Then GoTo ExitEnd

            If mStreamNumber = 0 Then
            ElseIf MIDIOutput1.StateStreamEx(mStreamNumber) = MIDISTATE_CLOSED Then
            Else
                Exit Do ' keep checking
            End If

            If cStreamName = "" Then
            ElseIf MIDIOutput1.StateStreamNameEx(cStreamName, mStreamNumber) = MIDISTATE_CLOSED Then ' find nearest stream
            'ElseIf MIDIOutput1.StateStreamNameEx(MIDISTATE_STARTED, mStreamNumber) = MIDISTATE_CLOSED Then ' find first started stream
            'ElseIf MIDIOutput1.StateStreamNameEx("", mStreamNumber) = MIDISTATE_CLOSED Then ' find first stream not closed
            Else
                Exit Do ' keep checking
            End If
            'MIDIOutput1.StreamNumber = mStreamNumber
            
            ' Wait for other streams too (optional).
            ' Wait based on other criteria (optional).
            
            isStreamWaitExit = True ' no more found
            Exit Do ' only run once
        Loop
        If isStreamWaitExit = True Then
            Exit Do  ' no more needed to check
        End If
    Loop
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub WaitSortStream( _
   ByVal mStreamNumber As Integer _
 , ByRef MIDIOutput1 As MIDIOutput)
    
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    '-------------------------------------------------------------------
    ' Wait until all related streams be sorted
    ' if object is enabled. Routine is polling for changes
    ' which is not efficient but sleep helps minimize cpu usage.
    ' The form and application will be accessible during loop so any objects
    ' should have already been disabled that are not supposed to run meanwhile.
    '-------------------------------------------------------------------
        
    Dim mCount As Integer
    Dim nMessageUBound As Long
    Dim mStateSortStreamPercent As Long
    
    ' Preserve passed data so not interfere with other functions
    'Dim backupelement As Integer
    'Call MidiStackPushCommon(backupelement, MIDIOutput1)
    Dim backupstreamnumber As Integer
    backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    ' (only use temporarily because not multhread-safe to use while in loop)
    
    MIDIOutput1.StreamNumber = mStreamNumber
    nMessageUBound = MIDIOutput1.StreamMessageUBound
    MIDIOutput1.StreamNumber = backupstreamnumber ' not needed anymore
    
    'If MainStreamOption <> MB_OPTIONOPENDEFAULT Then ' (optional)
    
    'MIDIOutput1.StreamNumber = mStreamNumber ' not multithread-safe
    'Do While MIDIOutput1.StateSortStreamEx(0) <> MIDISTATE_CLOSED ' alternative, not multithread-safe
    
    mCount = 0
    Do While MIDIOutput1.StateSortStreamEx(mStreamNumber) <> MIDISTATE_CLOSED
        If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
        
        Sleep MB_DOEVENTSPOLLING ' release resources enough so <5% cpu usage
        Call DoEventsOnce(False): If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
        'If MIDIOutput1.StreamNumber <> mStreamNumber Then Call Error3886 ' use if doevents screws it up
        'MIDIOutput1.MessageEventPending = False ' alternative
        
        If Int(mCount / 10) = mCount / 10 Then ' interval = 10 * MB_DOEVENTSPOLLING
            mStateSortStreamPercent = MIDIOutput1.StateSortStreamPercentEx(mStreamNumber) ' may sort logarithmically slower
            Me.Caption = "MFPlayer Example - Sorting - " _
             & Trim$(str$(nMessageUBound)) & " messages, " _
            & Trim$(str$(mStateSortStreamPercent)) & "%" _
            & IIf(mStateSortStreamPercent = 50, "+", "")
            mCount = 0 ' reset so not overflow
        End If
        mCount = mCount + 1
        
        'Exit Do
    Loop
    
ExitSection:

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Public Sub MidiStackPushCommon(ByRef backupelement As Integer, ByRef MIDIOutput1 As MIDIOutput)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close ' must not run until verified

    '-------------------------------------------------------------------
    ' Preserve passed data so not interfere with other functions
    '-------------------------------------------------------------------
    'Dim backupelement As Integer
    backupelement = MIDIOutput1.StackPush(0) ' zero for next available
    If backupelement = 0 Then _
     Err.Raise 1, , "PROGRAM ERROR 3875, forgot to pop somewhere before"
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Public Sub MidiStackPopCommon(ByRef backupelement As Integer, ByRef MIDIOutput1 As MIDIOutput)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    If MIDIOutput1.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close ' must not run until verified

    '-------------------------------------------------------------------
    ' Restore data
    ' last code so it is not affected by anything
    '-------------------------------------------------------------------
    If backupelement = 0 Then  ' nothing to restore
    ElseIf MIDIOutput1.StackPop(backupelement) = backupelement Then ' okay
    Else
        Err.Raise 1, , "PROGRAM ERROR 3876, something interrupted the previous push"
    End If

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Public Sub CommandStreamScanCopyMessage_ClickCommon( _
   ByVal mStreamNumberOld As Integer _
 , ByVal mStreamNumber As Integer _
 , Optional ByVal nMessageLBound As Long = -1 _
 , Optional ByVal nMessageUBound As Long = -1 _
 , Optional ByVal isSortOutOfOrder As Boolean = False _
 , Optional ByVal isFilterNonEnabled As Boolean = False _
 , Optional ByVal isFilterNotes As Boolean = False _
 , Optional ByVal isFilterDupl As Boolean = False)
    
    If gisEnd = True Then GoTo ExitEnd

    '-------------------------------------------------------------------
    ' Copy messages of a old stream into a new stream
    '-------------------------------------------------------------------
    
    'Dim mStreamNumber As Integer ' (comment out if passed as parameter already)
    Dim cStreamName As String
    Dim nStreamTimeStartRelativeToThisProcedure As Long
    Dim nStreamTimeStartRelativeToOpen As Long
    Dim isStreamWaitExit As Boolean

    Dim MIDIOutput1_MP(0 To MIDIMP_UBOUND) As Long ' always start from zero
    Dim nMP As Long
    
    Dim isSkip As Boolean
    Dim longtemp As Long
    Dim oldstream As Integer
    Dim oldmessage As Long
    Dim newstream As Integer
    Dim newmessage As Long
    Dim isError As Boolean
    
    Dim tempmessagestate As Integer
    Dim tempmessage As Integer
    
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative

    Do
        If gisEnd = True Then GoTo ExitEnd
        
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' Default to copy all of stream
        ''''''''''''''''''''''''''''''''''''''''''''''
        MIDIOutput1.StreamNumber = mStreamNumberOld
        If nMessageLBound = -1 Then nMessageLBound = 1 ' first message
        If nMessageUBound = -1 Then nMessageUBound = MIDIOutput1.StreamMessageUBound
        If nMessageUBound < 0 Then
            ' at least one message
            isError = True
        ElseIf nMessageUBound < nMessageLBound Then
            ' negative range not implemented
            isError = True
        ElseIf nMessageUBound > MIDIOutput1.StreamMessageUBound Then
            ' unknown message contents after ubound
            ' since array is not initialized nor cleared there
            isError = True
        End If
        If isError = True Then
            Err.Raise 1, , "PROGRAM ERROR 27864"
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''
        oldstream = mStreamNumberOld
        'oldmessage = ...  ' determine in loop
        newstream = mStreamNumber
        'newmessage = ...  ' not needed when queueing
    
#If 1 = 0 Then ' comment out to enable test
        isFilterNonEnabled = True
        isFilterNotes = True
        isFilterDupl = True
#End If
    
        If isFilterNonEnabled = True _
         And isFilterNotes = False _
         And isFilterDupl = False Then
            MIDIOutput1.StreamNumber = newstream
            If isSortOutOfOrder = True Then _
             MIDIOutput1.StreamMessageSortOutOfOrder = isSortOutOfOrder ' no sorting for speed
            Call MIDIOutput1.CopyStreamEx(oldstream, nMessageLBound, nMessageUBound, newstream, 1, 0)
        
        ElseIf isFilterNotes = True _
         And isFilterDupl = False Then
            MIDIOutput1.StreamNumber = newstream
            If isSortOutOfOrder = True Then _
             MIDIOutput1.StreamMessageSortOutOfOrder = isSortOutOfOrder ' no sorting for speed
            Call MIDIOutput1.CopyStreamEx(oldstream, nMessageLBound, nMessageUBound, newstream, 2, 0)
        
        ElseIf isFilterDupl = True Then
            Err.Raise 1, , "PROGRAM ERROR 2675, bad condition"
            'MIDIOutput1.StreamNumber = newstream
            'If isSortOutOfOrder = True Then _
            ' MIDIOutput1.StreamMessageSortOutOfOrder = isSortOutOfOrder ' no sorting for speed
            'Call MIDIOutput1.CopyStreamEx(oldstream, nMessageLBound, nMessageUBound, newstream, 3, 0) ' not yet implemented
            
            ' Alternative, but should only disable
            ' when done copying all pending streams.
            'Call MIDIOutput1.CopyStreamEx(oldstream, nMessageLBound, nMessageUBound, newstream, 2, 0)
            'Call FilterDuplicateMessages(newstream) ' slow
            'Call FilterDuplicateMessages(newstream, frommessage, tomessage) ' generates more output since not comparing with all
            
        Else
            ' All, as is
            MIDIOutput1.StreamNumber = newstream
            If isSortOutOfOrder = True Then _
             MIDIOutput1.StreamMessageSortOutOfOrder = isSortOutOfOrder ' no sorting for speed
            Call MIDIOutput1.CopyStreamEx(oldstream, nMessageLBound, nMessageUBound, newstream, 0, 0)
        End If
                
        ' Statistics
#If 1 = 0 Then ' comment out to enable test
        Dim isShowOldSorted As Boolean
        Dim isShowNewSorted As Boolean
        Dim isShowNewAsIs As Boolean
        'isShowOldSorted = True ' comment out if not needed
        isShowNewSorted = True ' comment out if not needed
        'isShowNewAsIs = True ' comment out if not needed
                
        ' Verify type of messages before or after copy.
        ' Easier to see if testing with a midi 0 file format (one track only).
        ' Easier to see since results will sort by message type.
        ' Results should be similar to message types before/after copy.
        
        Stop ' ready to show results, clear output window manually as needed
        ' (if this stops more than once, then the midi 1 file format was used)
        
        If isShowOldSorted = True _
         Or isShowNewSorted = True Then
            Dim artempdata(&HFFFF&, 2) As Long
            Dim tempcount As Long
            Dim tempfound As Boolean
            Dim tempelement As Long
            Dim tempmessagegroup As Integer
            'Dim tempmessage As Integer
            Dim tempdata1 As Integer
            Dim tempdata2 As Integer
            tempcount = 0
            If isShowOldSorted = True Then
                MIDIOutput1.StreamNumber = oldstream
                Debug.Print "old", "stream " & Trim$(str$(MIDIOutput1.StreamNumber)) _
                 , "-------------------------------------------------"
                nMessageLBound = nMessageLBound
                nMessageUBound = nMessageUBound
                'nMessageLBound = MIDIOutput1.StreamMessageLBound ' not applicable
                'nMessageUBound = MIDIOutput1.StreamMessageUBound ' not applicable
            ElseIf isShowNewSorted = True Then
                MIDIOutput1.StreamNumber = newstream
                Debug.Print "new", "stream " & Trim$(str$(MIDIOutput1.StreamNumber)) _
                 , "-------------------------------------------------"
                nMessageLBound = MIDIOutput1.StreamMessageLBound
                nMessageUBound = MIDIOutput1.StreamMessageUBound
            End If
            For longtemp = nMessageLBound To nMessageUBound
                MIDIOutput1.StreamMessageNumber = longtemp
                MIDIOutput1.ActionStream = MIDIOUT_READ
                tempmessagegroup = (MIDIOutput1.Message And MB_HIGHNIBBLE)
                If MIDIOutput1.MessageState <> MIDIMESSAGESTATE_ENABLED Then ' ignored anyway
                Else
                    tempcount = tempcount + 1
                    If (MIDIOutput1.Message And MB_HIGHNIBBLE) = CONTROLLER_CHANGE Then
                        tempelement = MIDIOutput1.Message * &H100& + MIDIOutput1.Data1
                        artempdata(tempelement, 1) = artempdata(tempelement, 1) + 1 ' number of duplicates
                        artempdata(tempelement, 2) = MIDIOutput1.Data2 ' most recent value
                    Else
                        tempelement = MIDIOutput1.Message * &H100& + 0
                        artempdata(tempelement, 1) = artempdata(tempelement, 1) + 1 ' number of duplicates
                        artempdata(tempelement, 2) = MIDIOutput1.Data1 ' most recent value
                    End If
                End If
            Next longtemp
            tempcount = 0
            tempfound = False
            For longtemp = 1 To &HFFFF&
                tempelement = longtemp
                If artempdata(tempelement, 1) > 0 Then
                    Debug.Print artempdata(tempelement, 1),
                    tempmessagegroup = ((tempelement / &H100) And MB_HIGHNIBBLE)
                    If (tempmessagegroup And MB_HIGHNIBBLE) = CONTROLLER_CHANGE Then
                        tempmessage = tempelement / &H100
                        tempdata1 = (tempelement And MB_LOWBYTE)
                        tempdata2 = artempdata(tempelement, 2)
                        Debug.Print Trim$(str$(tempmessage)) _
                         & "(" & Trim$(Hex$(tempmessage)) & ")" _
                         & "-" & Trim$(str$(tempdata1)) _
                         & "-" & Trim$(str$(tempdata2))
                    Else
                        tempmessage = tempelement / &H100
                        tempdata1 = artempdata(tempelement, 2)
                        Debug.Print Trim$(str$(tempmessage)) _
                         & "(" & Trim$(Hex$(tempmessage)) & ")" _
                         & "-" & Trim$(str$(tempdata1))
                    End If
                    tempcount = tempcount + 1
                    tempfound = True
                End If
                If Int(tempcount / 190) = tempcount / 190 And tempfound = True Then
                'If Int(longtemp / 190) = longtemp / 190 Then ' not applicable
                    Stop
                    tempfound = False ' not needed anymore
                    Debug.Print "pause" ' output window limitation <200 rows
                End If
            Next longtemp
        End If
                
        If isShowNewAsIs = True Then
            MIDIOutput1.StreamNumber = newstream
            Debug.Print "new as is", "-------------------------------------------------"
            For longtemp = MIDIOutput1.StreamMessageLBound To MIDIOutput1.StreamMessageUBound
                MIDIOutput1.StreamMessageNumber = longtemp
                MIDIOutput1.ActionStream = MIDIOUT_READ
                If MIDIOutput1.MessageState <> MIDIMESSAGESTATE_ENABLED Then
                Else
                    Debug.Print Trim$(str$(MIDIOutput1.Message)) _
                     & "(" & Trim$(Hex$(MIDIOutput1.Message)) & ")" _
                     & "-" & Trim$(str$(MIDIOutput1.Data1)) _
                     & "-" & Trim$(str$(MIDIOutput1.Data2))
                End If
                If Int(longtemp / 190) = longtemp / 190 Then
                    Stop
                    Debug.Print "pause" ' output window limitation <200 rows
                End If
            Next longtemp
        End If
#End If
                
        Exit Do  ' only run once
    Loop
    
ExitSection:

    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Public Sub FilterDuplicateMessages( _
   ByVal mStreamNumber _
 , Optional ByVal nMessageLBound As Long = -1 _
 , Optional ByVal nMessageUBound As Long = -1)
    
    If gisEnd = True Then GoTo ExitEnd
    
    '-------------------------------------------------------------------
    ' Filter duplicate message commands whose details
    ' are not critical which could slow down the program.
    ' The exact number of messages that will be disabled
    ' is not crucial. The log file if any only needs to know
    ' that it occurred. This will help prepare any midi output
    ' with as few messages as possible. Keeping like patches,
    ' banks, some controllers and no duplicates.
    ' (see 1.00.568)

    ' Fast by scanning original data only once.

    ' WARNING,
    ' Also disables messages that represent on/off
    ' parameters as triggers when change from zero to 127.
    ' Would have been better to keep the most recent
    ' triggered pair or comparable group. Would have been
    ' best to at least keep the min and max values and
    ' not just one value.
    
    ' WARNING,
    ' Only removing duplicate commands, not duplicate
    ' commands and values. Since duplicate commands
    ' may contain different values, assuming the last value
    ' was more important. Also assuming input sends data
    ' in correct order in first place, no extra nrpn after
    ' data entry, etc. otherwise if sent in wrong order,
    ' this program may utilitize the wrong data later.

    ' WARNING,
    ' Not checking for duplicates in sysex because too
    ' complicated and slow.
    '-------------------------------------------------------------------
        
    'Dim thisform As Object
    'Set thisform = ParentControls.Item(1).Parent
    'While Not (TypeOf thisform Is Form)
    '    Set thisform = thisform.Parent
    'Wend
    ' (see 1.00.607)
    
    Dim nMessageCount As Long
    
    Dim MIDIOutput1_MP(0 To MIDIMP_UBOUND) As Long ' always start from zero
    Dim nMP As Long
    Dim tempmessage As Integer
    Dim tempdata1 As Integer
    'Dim tempdata2 As Integer
    'Dim temptime As Long
    'Dim temptag As Long
    
    Dim tempmessagestate As Integer
    Dim templogonly As Boolean
    Dim tempmessagegroup As Integer
    Dim tempelement As Long
    Dim tempelement2 As Long
    Dim tempelementmessage As Long
    Dim tempelementdata As Long
    
    ReDim arMidiFFTracking(0 To &HFFFF&) As MidiFFTracking
    
    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    
    ' Default to copy all of streams
    MIDIOutput1.StreamNumber = mStreamNumber
    If nMessageLBound = -1 Then nMessageLBound = 1 ' first message
    If nMessageUBound = -1 Then nMessageUBound = MIDIOutput1.StreamMessageUBound
    If nMessageUBound > MIDIOutput1.StreamMessageUBound _
     And nMessageUBound <= MIDIOutput1.StreamMessageNumberMax Then
        ' illegal because unknown message contents after ubound
        ' since array is never initialized nor cleared
        Err.Raise 1, , "PROGRAM ERROR 27854"
    End If
    If nMessageUBound <= 1 Then GoTo ExitSection ' nothing to do
    
    ' Search
    MIDIOutput1.StreamNumber = mStreamNumber
    'If mStreamNumber <> 0 Then ... already checked
    For nMessageCount = nMessageUBound To nMessageLBound Step -1
        If gisEnd = True Then GoTo ExitEnd
        ' check array backwards with the latest data first
        
        MIDIOutput1.StreamMessageNumber = nMessageCount
        MIDIOutput1.ActionStream = MIDIOUT_READ
        
        nMP = MIDIOutput1.MessagePointer(MIDIOutput1_MP(0), UBound(MIDIOutput1_MP))
        tempmessagestate = MIDIOutput1_MP(MIDIMP_MESSAGESTATE)
        tempmessage = MIDIOutput1_MP(MIDIMP_MESSAGE)
        tempdata1 = MIDIOutput1_MP(MIDIMP_DATA1)
        'tempdata2 = MIDIOutput1_MP(MIDIMP_DATA2)
        'temptime = MIDIOutput1_MP(MIDIMP_TIME)
        'temptag = MIDIOutput1_MP(MIDIMP_MESSAGETAG)
        'tempmessagestate = MIDIOutput1.MessageState
        'tempmessage = MIDIOutput1.Message
        'tempdata1 = MIDIOutput1.Data1
        'tempdata2 = MIDIOutput1.Data2
        'temptime = MIDIOutput1.Time
        'temptag = MIDIOutput1.MessageTag
        'Note that no need to use all properties and slow the program down
        
        templogonly = MIDIOutput1.MessageLogOnly
        
        tempmessagegroup = (tempmessage And MB_HIGHNIBBLE)
        'tempmessagegroup = Int2(tempmessage, &H10, 0) ' alternative
        
        If tempmessagestate <> MIDIMESSAGESTATE_ENABLED Then
            ' already will be skipped later anyway
        
        ElseIf templogonly = True Then
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' no logs needed
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' (WARNING, assuming lots of messages would
            ' make stream too slow and too big ...)
            MIDIOutput1.MessageState = MIDIMESSAGESTATE_DISABLED
            'MIDIOutput1.StreamNumber = mStreamNumber
            'If mStreamNumber <> 0 Then ... not applicable
            'MIDIOutput1.StreamMessageNumber = i
            'MIDIOutput1.StreamMessageNumber = 0 ' append
            MIDIOutput1.ActionStream = MIDIOUT_QUEUE
        
        ElseIf tempmessagegroup = note_on Or tempmessagegroup = note_off Then
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' no notes needed
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' (WARNING, assuming lots of messages would
            ' make stream too slow and too big ...)
            MIDIOutput1.MessageState = MIDIMESSAGESTATE_DISABLED
            'MIDIOutput1.StreamNumber = mStreamNumber
            'If mStreamNumber <> 0 Then ... not applicable
            'MIDIOutput1.StreamMessageNumber = i
            'MIDIOutput1.StreamMessageNumber = 0 ' append
            MIDIOutput1.ActionStream = MIDIOUT_QUEUE
        
        ElseIf tempmessage = META _
         Then
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' no meta needed
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' (WARNING, assuming lots of messages would
            ' make stream too slow and too big ...)
            MIDIOutput1.MessageState = MIDIMESSAGESTATE_DISABLED
            'MIDIOutput1.StreamNumber = mStreamNumber
            'If mStreamNumber <> 0 Then ... not applicable
            'MIDIOutput1.StreamMessageNumber = i
            'MIDIOutput1.StreamMessageNumber = 0 ' append
            MIDIOutput1.ActionStream = MIDIOUT_QUEUE
                
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' groups more difficult since they use a running status
        ' e.g.
        ' program-bank
        ' nrpn, 63,62,6,38(&h26)
        ' rpn, 65,64,6,38(&h26)
        ' sysex
            
        ElseIf tempmessagegroup = CONTROLLER_CHANGE _
         And (tempdata1 = 0 Or tempdata1 = &H20) _
         Then
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' bank, on current channel (LSB and MSB)
            ' Keep one before and after most recent program change.
            ''''''''''''''''''''''''''''''''''''''''''''''
            tempelement = tempmessage * &H100& + tempdata1
            tempelement2 = (tempmessage + &H10) * &H100& + 0 ' related program change
            
            If arMidiFFTracking(tempelement).nDetectedMessageNumber = 0 Then
                ' Found one, most recent copy
                arMidiFFTracking(tempelement).nDetectedMessageNumber = nMessageCount
            
            ElseIf arMidiFFTracking(tempelement).nDetectedMessageNumber > _
             arMidiFFTracking(tempelement2).nDetectedMessageNumber Then
                ' Found one before program change
                arMidiFFTracking(tempelement).nDetectedMessageNumber = nMessageCount

            Else
                ' Disable duplicate
                MIDIOutput1.MessageState = MIDIMESSAGESTATE_DISABLED
                'MIDIOutput1.StreamNumber = mStreamNumber
                'If mStreamNumber <> 0 Then ... not applicable
                'MIDIOutput1.StreamMessageNumber = i
                'MIDIOutput1.StreamMessageNumber = 0 ' append
                MIDIOutput1.ActionStream = MIDIOUT_QUEUE
            End If
        
        ElseIf tempmessagegroup = program_change _
         Then
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' program change, on current channel
            ' Keep most recent copy,
            ' since this completes the running status
            ''''''''''''''''''''''''''''''''''''''''''''''
            tempelement = tempmessage * &H100& + 0
            
            If arMidiFFTracking(tempelement).nDetectedMessageNumber = 0 Then
                ' Found one, most recent copy
                arMidiFFTracking(tempelement).nDetectedMessageNumber = nMessageCount
            
            Else
                ' Disable duplicate
                MIDIOutput1.MessageState = MIDIMESSAGESTATE_DISABLED
                'MIDIOutput1.StreamNumber = mStreamNumber
                'If mStreamNumber <> 0 Then ... not applicable
                'MIDIOutput1.StreamMessageNumber = i
                'MIDIOutput1.StreamMessageNumber = 0 ' append
                MIDIOutput1.ActionStream = MIDIOUT_QUEUE
            End If
            
        ElseIf tempmessagegroup = CONTROLLER_CHANGE _
         And (tempdata1 = &H63 Or tempdata1 = &H62) _
         Then
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' nrpn, on current channel
            ' Keep one before and after most recent dataentry.
            ' Specifically, keep each unique nrpn before each dataentry.
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' Leave as is.
            ' not yet implemented
        
        ElseIf tempmessagegroup = CONTROLLER_CHANGE _
         And (tempdata1 = &H65 Or tempdata1 = &H64) _
         Then
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' rpn, on current channel
            ' Keep one before and after most recent dataentry.
            ' Specifically, keep each unique nrpn before each dataentry.
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' Leave as is.
            ' not yet implemented
        
        ElseIf tempmessagegroup = CONTROLLER_CHANGE _
         And (tempdata1 = 6 Or tempdata1 = &H26) _
         Then
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' dataentry, on current channel
            ' Keep most recent copy,
            ' since this completes the running status
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' Leave as is.
            ' not yet implemented
        
        ElseIf tempmessage = SYSEX Then
            ' Leave as is.
            ' Either sysex is,
            ' - not a simple dataset,
            ' - sysex is a running status with more than one data
            '   then can not remove duplicates easily anyway
            ' - it is bulk dump of something so no duplicates anyway
            ' - it is a damaged sysex we cant do anything about anyway
            
        Else
            ''''''''''''''''''''''''''''''''''''''''''''''
            ' Assume no duplicates for all other messages.
            ' Keep most recent copy.
            ' E.g.
            ' Continuous controllers (&HB0,&HD0,&HE0,etc.),
            ' on/off controllers (&HB0),
            ' global controllers (&HF1-&HFF), etc.
            ''''''''''''''''''''''''''''''''''''''''''''''
            If tempmessagegroup = CONTROLLER_CHANGE Then
                ' Controllers are identified by two bytes.
                tempelement = tempmessage * &H100& + tempdata1
            Else
                ' All others are identified by one byte.
                tempelement = tempmessage * &H100& + 0
            End If
            'tempelementmessage = Int(tempelement / &H100&) ' extract from high order byte
            'tempelementdata = (tempelement And MB_LOWBYTE)
            'tempelementmessage = (tempelement And MB_HIGHBYTE) ' not applicable
            
            If arMidiFFTracking(tempelement).nDetectedMessageNumber = 0 Then
                ' Found one, most recent copy
                arMidiFFTracking(tempelement).nDetectedMessageNumber = nMessageCount
            
            Else
                ' Disable duplicate
                MIDIOutput1.MessageState = MIDIMESSAGESTATE_DISABLED
                'MIDIOutput1.StreamNumber = mStreamNumber
                'If mStreamNumber <> 0 Then ... not applicable
                'MIDIOutput1.StreamMessageNumber = i
                'MIDIOutput1.StreamMessageNumber = 0 ' append
                MIDIOutput1.ActionStream = MIDIOUT_QUEUE
            End If
        End If

        'If nMessageCount = MB_LONGUBOUND Then Exit For
    Next nMessageCount

    ' Statistics
#If 1 = 0 Then ' comment out to enable test
    Dim longtemp As Long
    Dim isShowNewSorted As Boolean
    Dim isShowNewAsIs As Boolean
    isShowNewSorted = True ' comment out if not needed
    'isShowNewAsIs = True ' comment out if not needed
            
    ' Verify type of messages after filter.
    ' Easier to see with midi 0 or 1 file format.
    ' Easier to see since results will sort by message type.
    ' Results should be similar to message types before/after copy.
    ' Results should show no messages for some that are always filtered.
    ' Results should show fewer messages for some that are duplicates.
    ' Results should show same message amount for some that are not filtered.
    ' Results should show same message amount for midi 0 or midi 1 file format.
    
    Stop ' ready to show results, clear output window manually as needed
    
    If isShowNewSorted = True Then
        Dim artempdata(&HFFFF&, 2) As Long
        Dim tempcount As Long
        Dim tempfound As Boolean
        'Dim tempelement As Long
        'Dim tempmessagegroup As Integer
        'Dim tempmessage As Integer
        'Dim tempdata1 As Integer
        Dim tempdata2 As Integer
        tempcount = 0
        If isShowNewSorted = True Then
            MIDIOutput1.StreamNumber = mStreamNumber
            Debug.Print "filtered", "-------------------------------------------------"
            'Debug.Print "filtered", "stream " & Trim$(Str$(MIDIOutput1.StreamNumber)) ' not needed
            nMessageLBound = MIDIOutput1.StreamMessageLBound
            nMessageUBound = MIDIOutput1.StreamMessageUBound
        End If
        For longtemp = nMessageLBound To nMessageUBound
            MIDIOutput1.StreamMessageNumber = longtemp
            MIDIOutput1.ActionStream = MIDIOUT_READ
            tempmessagegroup = (MIDIOutput1.Message And MB_HIGHNIBBLE)
            If MIDIOutput1.MessageState <> MIDIMESSAGESTATE_ENABLED Then ' ignored anyway
            Else
                tempcount = tempcount + 1
                If (MIDIOutput1.Message And MB_HIGHNIBBLE) = CONTROLLER_CHANGE Then
                    tempelement = MIDIOutput1.Message * &H100& + MIDIOutput1.Data1
                    artempdata(tempelement, 1) = artempdata(tempelement, 1) + 1 ' number of duplicates
                    artempdata(tempelement, 2) = MIDIOutput1.Data2 ' most recent value
                Else
                    tempelement = MIDIOutput1.Message * &H100& + 0
                    artempdata(tempelement, 1) = artempdata(tempelement, 1) + 1 ' number of duplicates
                    artempdata(tempelement, 2) = MIDIOutput1.Data1 ' most recent value
                End If
            End If
        Next longtemp
        tempcount = 0
        tempfound = False
        For longtemp = 1 To &HFFFF&
            tempelement = longtemp
            If artempdata(tempelement, 1) > 0 Then
                Debug.Print artempdata(tempelement, 1),
                tempmessagegroup = ((tempelement / &H100) And MB_HIGHNIBBLE)
                If (tempmessagegroup And MB_HIGHNIBBLE) = CONTROLLER_CHANGE Then
                    tempmessage = tempelement / &H100
                    tempdata1 = (tempelement And MB_LOWBYTE)
                    tempdata2 = artempdata(tempelement, 2)
                    Debug.Print Trim$(str$(tempmessage)) _
                     & "(" & Trim$(Hex$(tempmessage)) & ")" _
                     & "-" & Trim$(str$(tempdata1)) _
                     & "-" & Trim$(str$(tempdata2))
                Else
                    tempmessage = tempelement / &H100
                    tempdata1 = artempdata(tempelement, 2)
                    Debug.Print Trim$(str$(tempmessage)) _
                     & "(" & Trim$(Hex$(tempmessage)) & ")" _
                     & "-" & Trim$(str$(tempdata1))
                End If
                tempcount = tempcount + 1
                tempfound = True
            End If
            If Int(tempcount / 190) = tempcount / 190 And tempfound = True Then
            'If Int(longtemp / 190) = longtemp / 190 Then ' not applicable
                Stop
                tempfound = False ' not needed anymore
                Debug.Print "pause" ' output window limitation <200 rows
            End If
        Next longtemp
    End If
            
    If isShowNewAsIs = True Then
        MIDIOutput1.StreamNumber = mStreamNumber
        Debug.Print "new as is", "-------------------------------------------------"
        For longtemp = MIDIOutput1.StreamMessageLBound To MIDIOutput1.StreamMessageUBound
            MIDIOutput1.StreamMessageNumber = longtemp
            MIDIOutput1.ActionStream = MIDIOUT_READ
            If MIDIOutput1.MessageState <> MIDIMESSAGESTATE_ENABLED Then
            Else
                Debug.Print Trim$(str$(MIDIOutput1.Message)) _
                 & "(" & Trim$(Hex$(MIDIOutput1.Message)) & ")" _
                 & "-" & Trim$(str$(MIDIOutput1.Data1)) _
                 & "-" & Trim$(str$(MIDIOutput1.Data2))
            End If
            If Int(longtemp / 190) = longtemp / 190 Then
                Stop
                Debug.Print "pause" ' output window limitation <200 rows
            End If
        Next longtemp
    End If
#End If

ExitSection:

    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Public Sub MsgBoxBug(ByVal isCheck As Boolean)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    If isCheck = True Then
        If AmbientUserMode() = True Then
            ' Stop output during debug but maintain tempo just as it
            ' usually plays when run from an executable.
            Debug.Print "No processing while dialog is shown in debug mode."
            'Call MIDIOutput1.MessageBoxWarning ' describes bug
            MIDIOutput1.MessageEventEnable = False
            Call StopStuckNote
        End If
    Else
        MIDIOutput1.MessageEventEnable = True ' restore
    End If
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Public Sub DoEventsOnce(Optional ByVal isRestrict As Boolean = False)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' Only need in emergencies if doevents interferes with procedures
    ' and can not get other alternative safety features to work.
    ' DoEvents in slow loops and procedures is normally convenient
    ' and practical, but unfortunately creates a multithreaded program.
    ' Affects accessible form objects, timers, and events.
    
    If isRestrict = True Then
        ' Prepare to restrict access while releasing resources.
        ' Only need to setup once, and should not turn on/off repeatedly.
        
        Frame2.Enabled = False ' notify to block form objects
        'LockWindowUpdate (Me.hWnd) ' alternative, but can not redraw or shutdown
        'gisCurrentDoEvents = True ' notify to block timers and events
        ' other alternatives, are harder to implement ...
        ' none of these are needed if the rest of the program handles multithreading
    End If
        
    ' Release resources
    DoEvents: If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
        
    If isRestrict = True Then
        ' Restore later after calling procedure is done.
    
    Else
        ' Restore after calling procedure is done.
        Frame2.Enabled = True ' restore
        'LockWindowUpdate (0) ' restore
        'gisCurrentDoEvents = False ' restore
    End If
    
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub TimerProgressbar_Timer()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    '-------------------------------------------------------------------
    ' Update indicators and progress.
    '-------------------------------------------------------------------

    Dim n As Integer
    Dim mGroupNumber As Integer
    Dim mStreamNumber As Integer
    Dim cText As String

    ' Preserve passed data so not interfere with other functions
    Dim backupelement As Integer
    Call MidiStackPushCommon(backupelement, MIDIOutput1)
    'Dim backupstreamnumber As Integer
    'backupstreamnumber = MIDIOutput1.StreamNumber ' alternative
    
    If gisCurrentQueue = True Then GoTo ExitSection ' only need when multithreading with doevents
    If gisCurrentFF = True Then GoTo ExitSection ' only need when multithreading with doevents
    
    ' Progress bar decay speed is affected by the
    ' interval of TimerProgressBar.
    'TimerProgressBar.Interval = 55 ' assuming already set small
    
    For n = 0 To NumVULoaded - 1
        ' Progress bar
        If VIndicator1(n).value > VSliderVuBarDecay Then
            ' Begin decay
            'If CheckBarDecay.Value <> 0 ' not applicable since always decay
            VIndicator1(n).value = VIndicator1(n).value - VSliderVuBarDecay
        
        Else
            ' Turn off
            VIndicator1(n).value = 0
        End If
        
        ' Progress bar peak
        If CheckPeakHold.value <> 0 _
         And Val(CommandIndicatorPeak(n).Tag) < VSliderVuPeakHold / TimerProgressBar.Interval _
         Then
            ' Hold peak while counting down
            CommandIndicatorPeak(n).Tag = Trim$(str$(Val(CommandIndicatorPeak(n).Tag) + 1))
        
        ElseIf CheckPeakDecay.value <> 0 _
         And VIndicatorPeak(n).value > VSliderVuPeakDecay Then
            ' Begin decay
            VIndicatorPeak(n).value = VIndicatorPeak(n).value - VSliderVuPeakDecay
        
        Else
            ' Turn off
            VIndicatorPeak(n).value = 0
            CommandIndicatorPeak(n).Visible = False ' not show if track empty
        End If
        
        ' Position peak
        ' formula based on proportions, X/MaxWidth=Value/Max, X=Left
        CommandIndicatorPeak(n).Left = VIndicatorPeak(n).Left _
         + (VIndicatorPeak(n).value * VIndicatorPeak(n).Width) _
         / VIndicatorPeak(n).max
    Next n

    ' Update queuetime
    ' Assume no update if already changing with scroll() or change().
    If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
        ' Midi format 0
        If MainStreamNumber <> 0 Then
            mStreamNumber = MainStreamNumber
            'MIDIOutput1.StreamNumber = MainStreamNumber ' alternative
            If MIDIOutput1.StateStreamEx(mStreamNumber) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(mStreamNumber) <> MIDISTATE_STARTED Then ' leave bar adjustable
            Else
                Call DisplayScrollBar
            End If
        Else
            Call ClearScrollBar
        End If

        ' Text of stream information
        If MainStreamNumber <> 0 Then
            MIDIOutput1.StreamNumber = MainStreamNumber
            n = MainStreamNumber - 1 ' 1-based scale, assuming indicator=stream-1
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf CheckTextIndicator.value = 0 Then
                If TextIndicator1(n).Visible <> False Then _
                 TextIndicator1(n).Visible = False ' not needed anymore
            Else
                cText = Trim$(str$(MIDIOutput1.StreamNumber))
                cText = cText & "S" & Trim$(str$(MIDIOutput1.StateStreamEx(0)))
                If MIDIOutput1.StateStreamEx(0) <> MIDISTATE_CLOSED Then
                    'cText = cText & ", N" & Trim$(MIDIOutput1.StreamName)
                    cText = cText & ", B" & Trim$(str$(MIDIOutput1.StreamMessageLBound))
                    cText = cText & "-" & Trim$(str$(MIDIOutput1.StreamMessageUBound))
                    cText = cText & ", Start" & Trim$(str$(MIDIOutput1.StreamTimeStartRelativeToOpen))
                    cText = cText & ", Current" & Trim$(str$(MIDIOutput1.StreamTimeCurrent))
                    If MIDIOutput1.StreamAutoClose = True Then
                        cText = cText & ", AClose" ' not open, not replay, not close
                    ElseIf MIDIOutput1.StreamAutoStop = True Then
                        cText = cText & ", AStop" ' not close, not open, not replay
                    ElseIf MIDIOutput1.StreamAutoReplay = True Then
                        cText = cText & ", AReplay" ' not close, not open
                    ElseIf MIDIOutput1.StreamAutoStart = True Then
                        cText = cText & ", AOpen" ' not close
                    End If
                    'cText = cText & "" & IIf(Trim$(Str$(MIDIOutput1.StreamTranspose)) <> 0, _
                    ' ", Transpose" & Trim$(Str$(MIDIOutput1.StreamTranspose)), "")
                    'cText = cText & "" & IIf(Trim$(Str$(MIDIOutput1.StreamTempoRate)) <> 100, _
                    ' ", Tempo" & Trim$(Str$(MIDIOutput1.StreamTempoRate)), "")
                
                    ' Individual messages in stream, not shown here.
                End If
                TextIndicator1(n).Text = cText
                If TextIndicator1(n).Visible <> True Then _
                 TextIndicator1(n).Visible = True
            End If
        End If
    
    Else
        ' Midi format 1
        mGroupNumber = UBound(MainStreamGroup, 1) ' last is master track
        'For mGroupNumber = 1 To UBound(MainStreamGroup, 1) ' not needed
        If mGroupNumber <> 0 Then
            mStreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            'MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER) ' alternative
            If MIDIOutput1.StateStreamEx(mStreamNumber) = MIDISTATE_CLOSED Then ' no stream
            ElseIf MIDIOutput1.StateStreamEx(mStreamNumber) <> MIDISTATE_STARTED Then ' leave bar adjustable
            Else
                Call DisplayScrollBar
            End If
        Else
            Call ClearScrollBar
        End If
    
        ' Text of stream information
        For mGroupNumber = 1 To UBound(MainStreamGroup, 1) - 2 ' no track zero or master track
        'For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
            MIDIOutput1.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
            n = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER) - 1 ' 1-based scale, assuming indicator=stream-1
            If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
            ElseIf CheckTextIndicator.value = 0 Then
                If TextIndicator1(n).Visible <> False Then _
                 TextIndicator1(n).Visible = False ' not needed anymore
            Else
                cText = Trim$(str$(MIDIOutput1.StreamNumber))
                cText = cText & "S" & Trim$(str$(MIDIOutput1.StateStreamEx(0)))
                If MIDIOutput1.StateStreamEx(0) <> MIDISTATE_CLOSED Then
                    'cText = cText & ", N" & Trim$(MIDIOutput1.StreamName)
                    cText = cText & ", B" & Trim$(str$(MIDIOutput1.StreamMessageLBound))
                    cText = cText & "-" & Trim$(str$(MIDIOutput1.StreamMessageUBound))
                    cText = cText & ", Start" & Trim$(str$(MIDIOutput1.StreamTimeStartRelativeToOpen))
                    cText = cText & ", Current" & Trim$(str$(MIDIOutput1.StreamTimeCurrent))
                    If MIDIOutput1.StreamAutoStart = True Then
                        cText = cText & ", AStart" ' not close
                    ElseIf MIDIOutput1.StreamAutoReplay = True Then
                        cText = cText & ", AReplay" ' not close, not open
                    ElseIf MIDIOutput1.StreamAutoStop = True Then
                        cText = cText & ", AStop" ' not close, not open, not replay
                    ElseIf MIDIOutput1.StreamAutoClose = True Then
                        cText = cText & ", AClose" ' not open, not replay, not close
                    End If
                    'cText = cText & "" & IIf(Trim$(Str$(MIDIOutput1.StreamTranspose)) <> 0, _
                    ' ", Transpose" & Trim$(Str$(MIDIOutput1.StreamTranspose)), "")
                    'cText = cText & "" & IIf(Trim$(Str$(MIDIOutput1.StreamTempoRate)) <> 100, _
                    ' ", Tempo" & Trim$(Str$(MIDIOutput1.StreamTempoRate)), "")
                
                    ' Individual messages in stream, not shown here.
                End If
                TextIndicator1(n).Text = cText
                If TextIndicator1(n).Visible <> True Then _
                 TextIndicator1(n).Visible = True
            End If
        Next mGroupNumber
    
    End If

ExitSection:

    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    ' last code so it is not affected by anything

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
    TimerProgressBar.Enabled = False ' timers should not be running anymore
End Sub

Private Sub TimerMidiError_Timer()
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    If MIDIOutput1.ErrorCode <> 0 Then
        ' Errors were not handled before.
        Call ShowMidiError(MIDIOutput1.ErrorCode, MIDIOutput1.ErrorMessage _
         , MIDIOutput1.ErrorMessageSource, MIDIOutput1.ErrorCount, True)
        MIDIOutput1.ErrorCode = 0 ' not needed anymore
    End If
    'MIDIOutput1.ErrorScheme = ... already set
    'MIDIOutput1.ErrorHalt = ... already set

    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
    TimerMidiError.Enabled = False ' timers should not be running anymore
End Sub

Private Sub TimerEnd_Timer()
    'If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    Unload Me ' shutdown again
End Sub

