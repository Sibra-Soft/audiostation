VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Plugins 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plugins"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2865
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Plugins.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView Listview_Plugins 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form_Plugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim filter As String
Dim fh As String

fh = Dir("bass*.dll")   ' find 1st file

Do While (fh <> "")
    Dim plug As Long
    plug = BASS_PluginLoad(fh, 0)   ' plugin loaded...
    If (plug) Then
        Dim pinfo As BASS_PLUGININFO
        pinfo = BASS_PluginGetInfo(plug) ' get plugin info to add to the file selector filter...
        Dim a As Long
        For a = 0 To pinfo.formatc - 1
            filter = filter & "|" & VBStrFromAnsiPtr(BASS_PluginGetInfoFormat(plug, a).Name) & " (" & VBStrFromAnsiPtr(BASS_PluginGetInfoFormat(plug, a).exts) & ")" & " - " & fh ' format description
            filter = filter & "|" & VBStrFromAnsiPtr(BASS_PluginGetInfoFormat(plug, a).exts)  ' extension filter
        Next a
        
        ' add plugin to the list
        Listview_Plugins.ListItems.Add , , fh
    End If
    fh = Dir()  ' get next file
Loop
End Sub
