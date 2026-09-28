VERSION 5.00
Begin VB.Form Form_Settings_Recorder 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-"
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
   Icon            =   "Form_Settings_Recorder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   Tag             =   "1041"
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " T(1018) "
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
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option_WaveFile 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Wave File"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   1440
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox Combox_Rate 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox Combox_InputDevice 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input device:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   540
         Width           =   1140
      End
   End
   Begin Audiostation.ButtonBig Button_Cancel 
      Height          =   390
      Left            =   3330
      TabIndex        =   8
      Tag             =   "1004"
      Top             =   2520
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   688
      Caption         =   "T(1037)"
      TextAlignment   =   0
   End
   Begin Audiostation.ButtonBig Button_Save 
      Height          =   390
      Left            =   2040
      TabIndex        =   9
      Tag             =   "1004"
      Top             =   2520
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   688
      Caption         =   "T(1017)"
      TextAlignment   =   0
   End
End
Attribute VB_Name = "Form_Settings_Recorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Init As Boolean
Private Sub Button_Cancel_Click()
Unload Me
End Sub

Private Sub Button_Save_Click()
Call Extensions.INIWrite("main", "RecordingDeviceId", Combox_InputDevice.ListIndex, ConfigFile)
Call Extensions.INIWrite("main", "RecordingRate", Combox_Rate.Text, ConfigFile)

Form_Main.DataInter.SendData "Done"

End
End Sub
Private Sub Form_Load()
Dim DefaultDevice As Integer
Dim SavedDevice As Integer

Call TranslateFormAndControls(Me)

With Combox_Rate
    .AddItem "96000"
    .AddItem "48000"
    .AddItem "44100"
    .AddItem "32000"
    .AddItem "22050"
    
    .Text = Extensions.INIRead("main", "RecordingRate", ConfigFile, "44100")
End With

' Get record devices
Dim Device As mdlAdioDevice
For Each Device In Form_Main.AdioCore.GetListOfDevices
    If Device.dInput And Device.dIsLoopback Then
        Combox_InputDevice.AddItem Device.dName
        Combox_InputDevice.ItemData(Combox_InputDevice.NewIndex) = Device.dId
    End If
Next

' Select the default or saved device
If Combox_InputDevice.ListCount > 0 Then: Combox_InputDevice.ListIndex = Extensions.INIRead("main", "RecordingDeviceId", ConfigFile, 0)

Init = False
End Sub
