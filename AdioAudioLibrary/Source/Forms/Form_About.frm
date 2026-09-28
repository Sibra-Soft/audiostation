VERSION 5.00
Begin VB.Form Form_About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Adio Audio Library"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4950
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Button_Close 
      Caption         =   "&Close"
      Height          =   420
      Left            =   1860
      TabIndex        =   2
      Top             =   3105
      Width           =   1230
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form_About.frx":000C
      Height          =   615
      Left            =   173
      TabIndex        =   1
      Top             =   2160
      Width           =   4605
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form_About.frx":009D
      Height          =   1095
      Left            =   173
      TabIndex        =   0
      Top             =   945
      Width           =   4605
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2235
      Picture         =   "Form_About.frx":0186
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "Form_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Button_Close_Click()
Unload Me
End Sub
