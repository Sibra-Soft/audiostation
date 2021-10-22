VERSION 5.00
Begin VB.Form Form_Track_Properties 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Track Properties"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Track_Properties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Audiostation.ButtonBig Button_Close 
      Height          =   390
      Left            =   4800
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "Close"
      TextAlignment   =   0
   End
   Begin VB.TextBox Textbox_Properties 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   50
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form_Track_Properties.frx":000C
      Top             =   50
      Width           =   5895
   End
End
Attribute VB_Name = "Form_Track_Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Close_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim BassTime As New BaseTime

Mp3Info.Filename = CurrentMediaFilename

With BassTime
    Textbox_Properties.Text = "Total duration: " & Format(.GetDuration(chan), "0.0") & " seconds / " & .GetTime(.GetDuration(chan)) & vbNewLine & _
    "Frequency: " & .GetFrequency(chan) & " Hz, " & .GetBits(chan) & " bits, " & .GetMode(chan) & vbNewLine & _
    "Bytes/s: " & .GetBytesPerSecond(chan) & vbNewLine & _
    "Kbp/s: " & Mp3Info.BitRate & vbNewLine & vbNewLine & _
    "Artist: " & Mp3Info.Artist & vbNewLine & _
    "Title: " & Mp3Info.Title & vbNewLine & _
    "Year: " & Mp3Info.Year
End With
End Sub

