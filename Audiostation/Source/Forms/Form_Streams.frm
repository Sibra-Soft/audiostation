VERSION 5.00
Begin VB.Form Form_Streams 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-"
   ClientHeight    =   6885
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Streams.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1046"
   Begin VB.ListBox ListBox_Streams 
      Height          =   6105
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin Audiostation.ButtonBig Button_Cancel 
      Height          =   390
      Left            =   2050
      TabIndex        =   1
      Tag             =   "1004"
      Top             =   6360
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   688
      Caption         =   "T(1037)"
      TextAlignment   =   0
   End
   Begin Audiostation.ButtonBig Button_OK 
      Height          =   390
      Left            =   945
      TabIndex        =   2
      Tag             =   "1004"
      Top             =   6360
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   688
      Caption         =   "T(1045)"
      TextAlignment   =   0
   End
End
Attribute VB_Name = "Form_Streams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StreamLinks As New Collection
Private Sub Button_Cancel_Click()
Unload Me
End Sub

Private Sub Button_OK_Click()
If ListBox_Streams.SelCount <= 0 Then: Exit Sub

StreamUrl = StreamLinks(ListBox_Streams.ListIndex + 1)
StreamName = ListBox_Streams.Text

Unload Me
End Sub

Private Sub Form_Load()
Dim StreamDb As String
Dim StreamItem As Variant

ListBox_Streams.Clear
StreamDb = Extensions.FileGetContents(App.path & "\streams.db")

For Each StreamItem In Split(StreamDb, vbNewLine)
    Dim StreamName, StreamLink As String
        
    StreamName = Trim(StrExt.SplitStr(CStr(StreamItem), ";", 0))
    StreamLink = Trim(StrExt.SplitStr(CStr(StreamItem), ";", 1))
    
    StreamLinks.Add StreamLink
    
    ListBox_Streams.AddItem StreamName
Next

Call TranslateFormAndControls(Me)
End Sub
Private Sub ListBox_Streams_DblClick()
Button_OK_Click
End Sub
