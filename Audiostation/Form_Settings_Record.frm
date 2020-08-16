VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_Settings_Record 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audio Record Settings"
   ClientHeight    =   2085
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6495
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
   ScaleHeight     =   2085
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1042"
   Begin VB.ComboBox cboSaveAs 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   900
      Width           =   3735
   End
   Begin VB.ComboBox cboAudioDevices 
      Height          =   315
      ItemData        =   "Form_Settings_Record.frx":000C
      Left            =   240
      List            =   "Form_Settings_Record.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   480
      Width           =   6015
   End
   Begin Audiostation.ButtonBig cmdCancel 
      Height          =   390
      Left            =   3000
      TabIndex        =   6
      Tag             =   "1001"
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Cancel"
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin Audiostation.ButtonBig cmdSave 
      Height          =   390
      Left            =   1800
      TabIndex        =   7
      Tag             =   "1029"
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Save"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSaveAudio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save audio file as:"
      Height          =   195
      Left            =   720
      TabIndex        =   10
      Top             =   960
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Audio Recording Device:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Tag             =   "1040"
      Top             =   240
      Width           =   2115
   End
End
Attribute VB_Name = "Form_Settings_Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GetRecordDevices()
Dim I As Integer
Dim deviceList As String
Dim Lines
Dim strPos As Integer
Dim startReadingDevices As Boolean

cboAudioDevices.Clear

deviceList = Extensions.FileGetContents(App.path & "\devices.txt")
Lines = Split(deviceList, vbNewLine)

For I = 0 To UBound(Lines)
    If startReadingDevices Then
        strPos = InStr(1, Lines(I), "]", vbTextCompare)
        
        If strPos > 0 Then
            cboAudioDevices.AddItem Replace(Trim(Mid(Lines(I), strPos + 1)), Chr(34), vbNullString)
        End If
    End If
    
    If InStr(1, Lines(I), "DirectShow audio devices", vbTextCompare) Then
        startReadingDevices = True
    End If
Next
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecordingDevice", cboAudioDevices.Text)
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecordingFileType", cboSaveAs.ListIndex)

Unload Me
End Sub
Private Sub Form_Load()
Dim SavedRecordingDevice As String

Call SetLanguage(Me)

cboAudioDevices.Clear
cboSaveAs.Clear

Call GetRecordDevices

' Add the audio types to the dropdown
cboSaveAs.AddItem "Microsoft WaveForm Audio (*.wav)"
cboSaveAs.AddItem "MPEG-1 Layer 3 (*.mp3)"

SavedRecordingDevice = Settings.ReadSetting("Sibra-Soft", "Audiostation", "RecordingDevice")
If Not SavedRecordingDevice = vbNullString And cboAudioDevices.ListCount > 0 Then
    cboAudioDevices.Text = Settings.ReadSetting("Sibra-Soft", "Audiostation", "RecordingDevice", vbNullString)
End If

cboSaveAs.ListIndex = Settings.ReadSetting("Sibra-Soft", "Audiostation", "RecordingFileType", 0)
End Sub

