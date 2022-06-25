VERSION 5.00
Begin VB.Form Form_Settings 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recorder Settings"
   ClientHeight    =   3045
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   Begin Recorder.ButtonBig Button_Save 
      Height          =   390
      Left            =   2115
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Save"
      TextAlignment   =   0
   End
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
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.OptionButton Option_Mp3File 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mp3 File"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option_WaveFile 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Wave File"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   1440
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox Combox_Rate 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox Combox_InputDevice 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input device:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   1140
      End
   End
   Begin Recorder.ButtonBig Button_Cancel 
      Height          =   390
      Left            =   3315
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Cancel"
      TextAlignment   =   0
   End
End
Attribute VB_Name = "Form_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim init As Boolean
Private Sub Button_Cancel_Click()
Unload Me
End Sub

Private Sub Button_Save_Click()
Dim index As Integer
Dim device As Long

index = Combox_InputDevice.ListIndex
device = Combox_InputDevice.ItemData(index)

Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecorderDeviceIndex", index)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecorderDevice", device)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecordingRate", Combox_Rate.text)

Unload Me
End Sub

Private Sub Form_Load()
Dim defaultDevice As Integer
Dim savedDevice As Integer

With Combox_Rate
    .AddItem "96000"
    .AddItem "48000"
    .AddItem "44100"
    .AddItem "32000"
    .AddItem "22050"
    
    .text = Settings.ReadSetting("Sibra-Soft", "Audiostation", "RecordingRate", "48000")
End With

' get list of WASAPI input devices
Dim c As Integer, i As Integer, di As BASS_WASAPI_DEVICEINFO

c = 0
i = 0

While BASS_WASAPI_GetDeviceInfo(c, di)
    If ((di.flags And BASS_DEVICE_INPUT) = BASS_DEVICE_INPUT And (di.flags And BASS_DEVICE_ENABLED) = BASS_DEVICE_ENABLED) Then ' it's an enabled input device
        With Combox_InputDevice
            .AddItem VBStrFromAnsiPtr(di.name)
            i = .ListCount - 1
            .ItemData(i) = c    ' retain device # for later
        End With
        If ((di.flags And BASS_DEVICE_DEFAULT) = BASS_DEVICE_DEFAULT) Then ' it's the default
            indev = c
            ' initialize device
            init = True
            defaultDevice = i
        End If
    End If
    c = c + 1
Wend

' Select the default or saved device
savedDevice = Settings.ReadSetting("Sibra-Soft", "Audiostation", "RecorderDeviceIndex", 0)
If savedDevice = 0 Then
    Combox_InputDevice.ListIndex = defaultDevice
Else
    Combox_InputDevice.ListIndex = Settings.ReadSetting("Sibra-Soft", "Audiostation", "RecorderDeviceIndex", 0)
End If

init = False

If (indev = -1) Then Call Error_("Can't find any WASAPI input devices")
End Sub
