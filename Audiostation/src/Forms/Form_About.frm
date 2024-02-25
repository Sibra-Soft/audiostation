VERSION 5.00
Begin VB.Form Form_About 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-"
   ClientHeight    =   5430
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_About.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1013"
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6135
      TabIndex        =   11
      Top             =   0
      Width           =   6135
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Form_About.frx":000C
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audiostation"
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
         Left            =   720
         TabIndex        =   13
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The old school media player"
         Height          =   195
         Left            =   720
         TabIndex        =   12
         Top             =   340
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   6135
      TabIndex        =   3
      Top             =   2640
      Width           =   6135
      Begin Audiostation.Hyperlink Hyperlink_Website 
         Height          =   195
         Left            =   1680
         TabIndex        =   15
         Top             =   1680
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   344
         Caption         =   "audiostation.org"
         URL             =   "https://www.audiostation.org"
         BackColor       =   16777215
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visit our website:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form_About.frx":031F
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   400
         Width           =   5415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright � Sibra-Soft 2009 - 2023"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1260
         Width           =   3090
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alex van den Berg"
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
         Left            =   3480
         TabIndex        =   7
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application designer and programmer:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   3315
      End
      Begin VB.Label lbl_version 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
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
         Left            =   1600
         TabIndex        =   5
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current version:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1425
      End
   End
   Begin Audiostation.ButtonBig cmdClose 
      Height          =   390
      Left            =   3000
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "T(1025)"
      TextAlignment   =   0
   End
   Begin Audiostation.ButtonBig cmdThanks 
      Height          =   390
      Left            =   1800
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Thanks"
      TextAlignment   =   0
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form_About.frx":03B0
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Audiostation is a program to play music files. The program supports all common media files like (*.mp3, *.wav, *.mid, etc) "
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
End
Attribute VB_Name = "Form_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdThanks_Click()
MsgBox "Thanks to:" & vbNewLine & _
"- Pieter van den Broek" & vbNewLine & _
"- N.Heine (Creative Bug Finder)" & vbNewLine & vbNewLine & _
"- BroeX Audiovisuele Producties" & vbNewLine & _
"- Voyetra Software" & vbNewLine & _
"- High Voltage SID Collection" & vbNewLine & _
"- Virtual Midi Synth" & vbNewLine & _
"- SID Player" & vbNewLine & _
"- Rogerd Pack (Screen Capture Recorder)" & vbNewLine & _
"- Mabry Software, Inc. (Midi Device Controller)" & vbNewLine & _
"- BeepBox.co" & vbNewLine & _
"- ffmpeg.org", vbInformation, "Thanks"
End Sub

Private Sub Form_Load()
lbl_version.Caption = App.Major & "." & App.Minor & " Build: " & App.Revision

Call TranslateFormAndControls(Me)
End Sub

