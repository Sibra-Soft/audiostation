VERSION 5.00
Begin VB.Form Form_Settings 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audiostation Settings"
   ClientHeight    =   4965
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Settings_Record.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1042"
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Settings "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin Audiostation.Hyperlink Hyperlink_ChangeRecorderSettings 
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2005
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   344
         Caption         =   "Change audio recording settings"
         ClickResponse   =   1
         BackColor       =   12632256
         ColorNormal     =   16711680
         ColorHot        =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox Combox_MidiDevice 
         Height          =   315
         ItemData        =   "Form_Settings_Record.frx":000C
         Left            =   3120
         List            =   "Form_Settings_Record.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show the spectrum analyzer"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show the application CD animation"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.ComboBox Combox_Language 
         Height          =   315
         ItemData        =   "Form_Settings_Record.frx":0035
         Left            =   3120
         List            =   "Form_Settings_Record.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   800
         Width           =   3135
      End
      Begin VB.ComboBox Combox_PlaybackDevice 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3600
         Width           =   6135
      End
      Begin VB.ComboBox Combox_CDRomDrive 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   420
         Width           =   3135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Midi playback device:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1260
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   855
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audio plackback device:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CD-Rom Drive:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin Audiostation.ButtonBig Button_Cancel 
      Height          =   390
      Left            =   3405
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Cancel"
      TextAlignment   =   0
   End
   Begin Audiostation.ButtonBig Button_Save 
      Height          =   390
      Left            =   2175
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Save"
      TextAlignment   =   0
   End
End
Attribute VB_Name = "Form_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Cancel_Click()
Unload Me
End Sub

Private Sub Button_Save_Click()
Dim PlaybackDevice As Long

PlaybackDevice = Combox_PlaybackDevice.ItemData(Combox_PlaybackDevice.ListIndex)

Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CDDeviceId", Combox_CDRomDrive.ListIndex)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CDDeviceLetter", left(Combox_CDRomDrive.text, 2))
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "PlaybackDevice", PlaybackDevice)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "PlaybackDeviceId", Combox_PlaybackDevice.ListIndex)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "Langauge", Combox_Language.text)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "MidiPlaybackDeviceId", Combox_MidiDevice.ListIndex)

Call BASS_SetDevice(PlaybackDevice)

Unload Me
End Sub

Private Sub Form_Load()
Dim a As Long, n As Long
Dim cdi As BASS_CD_INFO

' Get CD devices
a = 0
While (a < MAXDRIVES And BASS_CD_GetInfo(a, cdi) <> 0)
    Combox_CDRomDrive.AddItem Chr$(65 + cdi.letter) & ": " & VBStrFromAnsiPtr(cdi.vendor) & " " & VBStrFromAnsiPtr(cdi.product) & " " & VBStrFromAnsiPtr(cdi.rev)    ' "letter: description"
    a = a + 1
Wend

If Combox_CDRomDrive.ListCount = 0 Then
    Combox_CDRomDrive.Enabled = False
    Combox_CDRomDrive.BackColor = vbButtonFace
    Label1.Enabled = False
End If

' Get playback devices
Dim C As Integer
Dim I As BASS_DEVICEINFO

C = 1      ' device 1 = 1st real device

While BASS_GetDeviceInfo(C, I)
    If (I.flags And BASS_DEVICE_ENABLED) Then  ' enabled, so add it...
        Combox_PlaybackDevice.AddItem VBStrFromAnsiPtr(I.name)
        Combox_PlaybackDevice.ItemData(Combox_PlaybackDevice.NewIndex) = C    'store device #
    End If
    C = C + 1
Wend

' Get MIDI devices
Combox_MidiDevice.Clear
Dim d As Integer
For d = 0 To Form_Midi.OutputDevCombo.ListCount
    If Form_Midi.OutputDevCombo.List(d) <> vbNullString Then
        Combox_MidiDevice.AddItem Form_Midi.OutputDevCombo.List(d)
    End If
Next

If Combox_PlaybackDevice.ListCount > 0 Then Combox_PlaybackDevice.ListIndex = Settings.ReadSetting("Sibra-Soft", "Audiostation", "PlaybackDeviceId", 0)

Combox_CDRomDrive.ListIndex = Settings.ReadSetting("Sibra-Soft", "Audiostation", "CDDeviceId", 0)
Combox_MidiDevice.ListIndex = Settings.ReadSetting("Sibra-Soft", "Audiostation", "MidiPlaybackDeviceId", 0)
Combox_Language.text = LCase(Settings.ReadSetting("Sibra-Soft", "Audiostation", "Langauge", "English"))
End Sub
Private Sub Hyperlink_ChangeRecorderSettings_Click()
Call AudioStaRecorder.Settings
End Sub
