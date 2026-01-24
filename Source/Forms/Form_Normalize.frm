VERSION 5.00
Begin VB.Form Form_Settings_AutoVolume 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Normalize"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Normalize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer_ControlActivator 
      Interval        =   10
      Left            =   2880
      Top             =   2640
   End
   Begin VB.OptionButton Option_ToTargetLevel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "To target level..."
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   370
      Width           =   2895
   End
   Begin VB.OptionButton Option_HighestPeak 
      BackColor       =   &H00C0C0C0&
      Caption         =   "To highest peak in sound"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1680
      ScaleHeight     =   240
      ScaleWidth      =   795
      TabIndex        =   10
      Top             =   2280
      Width           =   855
      Begin VB.TextBox Textbox_NormalizeAbove 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   10
         TabIndex        =   11
         Text            =   "99"
         Top             =   5
         Width           =   600
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   5
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1680
      ScaleHeight     =   240
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   1920
      Width           =   855
      Begin VB.TextBox Textbox_NormalizeBelow 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   10
         TabIndex        =   8
         Text            =   "85"
         Top             =   5
         Width           =   600
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   5
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1680
      ScaleHeight     =   240
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   840
      Width           =   855
      Begin VB.TextBox Textbox_NormalizeTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   10
         TabIndex        =   5
         Text            =   "98"
         Top             =   5
         Width           =   600
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   5
         Width           =   615
      End
   End
   Begin Audiostation.ButtonBig Button_Cancel 
      Height          =   390
      Left            =   1750
      TabIndex        =   15
      Tag             =   "1004"
      Top             =   3120
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   688
      Caption         =   "T(1037)"
      TextAlignment   =   0
   End
   Begin Audiostation.ButtonBig Button_OK 
      Height          =   390
      Left            =   500
      TabIndex        =   16
      Tag             =   "1004"
      Top             =   3120
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   688
      Caption         =   "T(1017)"
      TextAlignment   =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Normalize to"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "only if Peak Level is"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "below"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1950
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "or above"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   2310
      Width           =   855
   End
End
Attribute VB_Name = "Form_Settings_AutoVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Cancel_Click()
Unload Me
End Sub

Private Sub Button_OK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call TranslateFormAndControls(Me)
End Sub

Private Sub Timer_ControlActivator_Timer()
If Option_ToTargetLevel.Value Then
    Textbox_NormalizeAbove.Enabled = True
    Textbox_NormalizeBelow.Enabled = True
    Textbox_NormalizeTarget.Enabled = True
Else
    Textbox_NormalizeAbove.Enabled = False
    Textbox_NormalizeBelow.Enabled = False
    Textbox_NormalizeTarget.Enabled = False
End If
End Sub
