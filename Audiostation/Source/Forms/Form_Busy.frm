VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_Busy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-"
   ClientHeight    =   285
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6675
   ControlBox      =   0   'False
   Icon            =   "Form_Busy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1053"
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "Form_Busy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call TranslateFormAndControls(Me)
End Sub
