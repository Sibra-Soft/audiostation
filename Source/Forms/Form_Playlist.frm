VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form_Playlist 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "--"
   ClientHeight    =   8085
   ClientLeft      =   2565
   ClientTop       =   1800
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Playlist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1027"
   Begin VB.Timer Timer_Main 
      Interval        =   10
      Left            =   11400
      Top             =   3360
   End
   Begin Audiostation.ButtonBig Button_Save 
      Height          =   390
      Left            =   10200
      TabIndex        =   1
      Top             =   6720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
      Caption         =   "T(1017)"
      TextAlignment   =   0
   End
   Begin MSComctlLib.ListView Listview_Playlist 
      Height          =   6735
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11880
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1392
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   14658
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Width           =   1429
      EndProperty
   End
   Begin VB.FileListBox DirToPlaylist 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   4200
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combox_FilterType 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   8655
   End
   Begin Audiostation.ButtonBig Button_Close 
      Height          =   390
      Left            =   10200
      TabIndex        =   0
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
      Caption         =   "T(1025)"
      TextAlignment   =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   11400
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox ButtonBar 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   0
      Left            =   10080
      ScaleHeight     =   2535
      ScaleWidth      =   1935
      TabIndex        =   4
      Top             =   840
      Width           =   1935
      Begin Audiostation.ButtonBig Button_Delete 
         Height          =   390
         Left            =   120
         TabIndex        =   5
         Tag             =   "1004"
         Top             =   1440
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         Caption         =   "T(1023)"
         TextAlignment   =   0
      End
      Begin Audiostation.ButtonBig Button_ClearAll 
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Tag             =   "1005"
         Top             =   0
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         Caption         =   "T(1020)"
         TextAlignment   =   0
      End
      Begin Audiostation.ButtonBig Button_OpenPlaylist 
         Height          =   390
         Left            =   120
         TabIndex        =   7
         Tag             =   "1009"
         Top             =   960
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         Caption         =   "T(1022)"
         TextAlignment   =   0
      End
      Begin Audiostation.ButtonBig Button_SavePlaylist 
         Height          =   390
         Left            =   120
         TabIndex        =   8
         Tag             =   "1008"
         Top             =   480
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         Caption         =   "T(1021)"
         TextAlignment   =   0
      End
      Begin Audiostation.ButtonBig Button_Options 
         Height          =   390
         Left            =   120
         TabIndex        =   9
         Tag             =   "1054"
         Top             =   2040
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         Caption         =   "T(1024)"
         TextAlignment   =   0
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   9975
   End
   Begin VB.Label Label_PlaylistDetails 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 File(s) in this playlist, total duration 00:00:00"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   7740
      Width           =   4020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Playlist:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1155
   End
   Begin VB.Menu mnupopup 
      Caption         =   "- POPUP -"
      Begin VB.Menu menu_Popup_AddFiles 
         Caption         =   "&Add File(s)"
         HelpContextID   =   1028
      End
      Begin VB.Menu menu_Popup_AddDirectory 
         Caption         =   "&Add Directory"
         HelpContextID   =   1029
      End
      Begin VB.Menu space01 
         Caption         =   "-"
      End
      Begin VB.Menu menu_Popup_ShufflePlaylist 
         Caption         =   "&Shuffle Playlist"
         HelpContextID   =   1030
      End
      Begin VB.Menu menu_Popup_HtmlPlaylist 
         Caption         =   "&Generate HTML playlist"
         HelpContextID   =   1031
      End
   End
End
Attribute VB_Name = "Form_Playlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim InitDone As Boolean
Dim DoneFitToSize As Boolean
Dim HasPlayingItem As Boolean

Public Enum enumFormTypes
    [MidiPlayer]
    [Mp3Player]
End Enum

Public FormType As enumFormTypes
Private Sub FitListviewToSize(Optional Force As Boolean)
Dim Offset As Integer

Offset = 250

'Do nothing if it's already been done
If DoneFitToSize = True Then: Exit Sub

'Check if the scrollbar is visible
If Listview_Playlist.ListItems.Count > 31 Or Force = True Then
    Form_Playlist.Listview_Playlist.ColumnHeaders(1).Width = Form_Playlist.Listview_Playlist.ColumnHeaders(1).Width - Offset
    
    DoneFitToSize = True
End If
End Sub
Private Function ComboxboxToCommondialogFilter() As String
Dim I As Integer
Dim FileType As String
Dim Description As String
Dim FilterString As New clsStringBuilder

For I = 0 To Combox_FilterType.ListCount - 1
    FileType = StrExt.Between("(", ")", Combox_FilterType.List(I))
    Description = StrExt.SplitStr(Combox_FilterType.List(I), "(", 0)
    
    FilterString.Append Description & "(" & FileType & ")|" & FileType & "|"
Next

ComboxboxToCommondialogFilter = FilterString.toString
End Function
Private Sub SavePlaylistAsHTML(FileName As String)
Dim I As Integer
Dim TemplateContent As String
Dim table As New clsStringBuilder

'Open the file with the template content
TemplateContent = Extensions.FileGetContents(App.Path & "\templates\html_playlist.tpl")

'Create the table
table.AppendNL "<table style='width:100%;' border='2px' >"
table.AppendNL "<thead>"
    table.AppendNL "<tr>"
        table.AppendNL "<td><strong>Nr.</strong></td>"
        table.AppendNL "<td><strong>Filename</strong></td>"
        table.AppendNL "<td><strong>Duration</strong></td>"
    table.AppendNL "</tr>"
table.AppendNL "</thead>"

For I = 1 To Listview_Playlist.ListItems.Count
    table.AppendNL "<tr>"
        table.AppendNL "<td>" & Listview_Playlist.ListItems(I).Text & "</td>"
        table.AppendNL "<td>" & Listview_Playlist.ListItems(I).SubItems(1) & "</td>"
        table.AppendNL "<td>" & Listview_Playlist.ListItems(I).SubItems(2) & "</td>"
    table.AppendNL "</tr>"
Next

table.AppendNL "</table>"

'Replace the template variable with the created table
TemplateContent = Replace(TemplateContent, "{PlaylistTable}", table.toString)

Call Extensions.FilePutContents(FileName, TemplateContent)
End Sub
Public Sub AddToPlaylist(File As String)
Dim Files
Dim FilesToAddCount, FilesAddedCount As Integer
Dim I As Integer
Dim FirstIndex As Integer
Dim AddingMoreThanOne As Boolean
Dim FileNotFoundCount As Integer

'On Error GoTo ErrorHandler
Files = Split(File, vbNewLine)
FilesToAddCount = UBound(Files)

If FilesToAddCount > 0 Then: FirstIndex = 0
If FilesToAddCount > 1 Then: AddingMoreThanOne = True

'More than one file, so load the busy form
If AddingMoreThanOne = True Then
    Screen.MousePointer = vbHourglass
    
    Form_Busy.ProgressBar.Value = 0
    Form_Busy.ProgressBar.Max = FilesToAddCount + 1
    Form_Busy.Show , Me
End If

Debug.Print "Playlist: (Adding files to playlist)"

'Add files to the playlist
For I = FirstIndex To FilesToAddCount
    'Get current file to process
    File = Files(I)
    
    If Dir(File, vbDirectory) = vbNullString Then
        Debug.Print "File not found: " & File
        FileNotFoundCount = FileNotFoundCount + 1
    Else
        If FormType = Mp3Player Then: Form_Main.AdioMediaPlaylist.AddFile File
        If FormType = MidiPlayer Then: Form_Main.AdioMidiPlaylist.AddFile File
        
        If AddingMoreThanOne = True Then
            DoEvents
            Form_Busy.ProgressBar.Value = Form_Busy.ProgressBar.Value + 1
        End If
    End If
    
    FilesAddedCount = FilesAddedCount + 1
Next

'Check if all files are found
If FileNotFoundCount > 0 Then MsgBox StrExt.format(GetTranslation(1036), FileNotFoundCount), vbExclamation

Debug.Print "Playlist: (Done adding " & FilesAddedCount & " file(s), " & FileNotFoundCount & " file(s) not found while adding)"

'Restore default
Screen.MousePointer = vbDefault
Unload Form_Busy

Call PopulatePlaylist

Exit Sub

ErrorHandler:
Select Case Err.Number
    Case 0
    Case 35602
    Case Else: Debug.Print Err.Number & " - " & Err.Description
End Select
End Sub
Public Function PopulatePlaylist()
Dim PlaylistItem As mdlAdioPlaylistItem
Dim LstItem As ListItem

Listview_Playlist.ListItems.Clear

Select Case FormType
    Case enumFormTypes.MidiPlayer
        For Each PlaylistItem In Form_Main.AdioMidiPlaylist.GetList
            Set LstItem = Listview_Playlist.ListItems.Add(, , format(PlaylistItem.nR, "00"))
                LstItem.SubItems(1) = PlaylistItem.LocalFile
                LstItem.SubItems(2) = PlaylistItem.RuntimeString
            
            If PlaylistItem.nR = CurrentMidiPlayerTrackNr Then
                LstItem.Bold = True
                LstItem.ListSubItems(1).Bold = True
                LstItem.ListSubItems(2).Bold = True
            End If
        Next
        
    Case enumFormTypes.Mp3Player
        For Each PlaylistItem In Form_Main.AdioMediaPlaylist.GetList
            Set LstItem = Listview_Playlist.ListItems.Add(, , format(PlaylistItem.nR, "00"))
                LstItem.SubItems(1) = PlaylistItem.LocalFile
                LstItem.SubItems(2) = PlaylistItem.RuntimeString
            
            If PlaylistItem.nR = CurrentMediaPlayerTrackNr Then
                LstItem.Bold = True
                LstItem.ListSubItems(1).Bold = True
                LstItem.ListSubItems(2).Bold = True
            End If
        Next
End Select

Call FitListviewToSize
End Function
Public Function LoadFileTypes()
Select Case FormType
    Case enumFormTypes.MidiPlayer
        Combox_FilterType.AddItem GetTranslation(1032) & " (*.mid;*.midi;*.kar;*.mus;*.sid)"
        Combox_FilterType.AddItem "MIDI File (*.mid)"
        Combox_FilterType.AddItem "MIDI File (*.midi)"
        Combox_FilterType.AddItem "Karaoke File (*.kar)"
        Combox_FilterType.AddItem "Sibra-Soft Beep Symphony (*.mus)"
        Combox_FilterType.AddItem "Commodore64 Sound File (*.sid)"
        Combox_FilterType.AddItem "Musical Instrument Digital Interface (*.rmi)"
        Combox_FilterType.AddItem GetTranslation(1066) & " (*.*)"
        
        Combox_FilterType.ListIndex = 0 'Select the first item
        
    Case enumFormTypes.Mp3Player
        Combox_FilterType.AddItem GetTranslation(1032) & " (*.mp3;*.mp2;*.m4a;*.cda;*.wav;*.aif;*.wma;)"
        Combox_FilterType.AddItem "MPEG-1 Layer 3 (*.mp3)"
        Combox_FilterType.AddItem "MPEG-1 Layer 2 (*.mp2)"
        Combox_FilterType.AddItem "MPEG-4 Layer 4 Audio (*.m4a)"
        Combox_FilterType.AddItem "CD Audio (*.cda)"
        Combox_FilterType.AddItem "Microsoft WaveForm Audio (*.wav)"
        Combox_FilterType.AddItem "Audio Interchange File (*.aif)"
        Combox_FilterType.AddItem "Advanced Audio Coding (*.aac)"
        Combox_FilterType.AddItem "Windows Media Audio (*.wma)"
        Combox_FilterType.AddItem "Ogg Vorbis Audio File (*.ogg)"
        Combox_FilterType.AddItem "Free Lossless Audio Codec (*.flac)"
        Combox_FilterType.AddItem GetTranslation(1066) & " (*.*)"
        
        Combox_FilterType.ListIndex = 0 'Select the first item
End Select
End Function

Private Sub Button_Close_Click()
Unload Me
End Sub

Private Sub Button_Delete_Click()
Select Case FormType
    Case enumFormTypes.MidiPlayer
        Form_Main.AdioMidiPlaylist.Remove Listview_Playlist.SelectedItem
        
    Case enumFormTypes.Mp3Player
        Form_Main.AdioMediaPlaylist.Remove Listview_Playlist.SelectedItem
End Select

Listview_Playlist.ListItems.Remove Listview_Playlist.SelectedItem.Index
End Sub

Private Sub Button_OpenPlaylist_Click()
On Error GoTo ErrorHandler

Debug.Print "Playlist: (ShowOpenPlaylistDialog)"

With CommonDialog
    .CancelError = True
    .InitDir = Extensions.INIRead("main", "LastLocation", ConfigFile, App.Path)
    .DialogTitle = GetTranslation(1018)
    .Filter = "Audiostation Playlist (*.apl)|*.apl|" & GetTranslation(1019) & " (.m3u)|*.m3u|ShoutCast Playlist (*.pls)|*.pls|Windows Media Player Playlist (*.wpl)|*.wpl"
    .ShowOpen
    
    If .FilterIndex = 1 Then: Call Form_Main.AdioMediaPlaylist.LoadPlaylist(.FileName, PLAYLIST_APL)
    If .FilterIndex = 2 Then: Call Form_Main.AdioMediaPlaylist.LoadPlaylist(.FileName, PLAYLIST_M3U)
    If .FilterIndex = 3 Then: Call Form_Main.AdioMediaPlaylist.LoadPlaylist(.FileName, PLAYLIST_PLS)
    If .FilterIndex = 4 Then: Call Form_Main.AdioMediaPlaylist.LoadPlaylist(.FileName, PLAYLIST_WPL)

    Call PopulatePlaylist
    Call FitListviewToSize
    
    Debug.Print "Playlist: (OpenPlaylist: " & .FileName & ")"
End With

Call Extensions.INIWrite("main", "LastLocation", Extensions.GetPathFromFilename(CommonDialog.FileName), ConfigFile)
Exit Sub

ErrorHandler:
Select Case Err.Number
    Case 0
    Case Else: Debug.Print Err.Description
End Select
End Sub

Private Sub Button_Options_Click()
PopupMenu mnupopup
End Sub

Private Sub Button_Save_Click()
Unload Me
End Sub

Private Sub Button_SavePlaylist_Click()
On Error GoTo ErrorHandler
With CommonDialog
    .CancelError = True
    .FileName = App.Path
    .DialogTitle = GetTranslation(1017)
    .Filter = "Audiostation Playlist (*.apl)|*.apl|" & GetTranslation(1019) & " (.m3u)|*.m3u|ShoutCast Playlist (*.pls)|*.pls|Windows Media Player Playlist (*.wpl)|*.wpl"
    .ShowSave

    If Right(LCase(.FileName), 3) = "apl" Then: Call Form_Main.AdioMediaPlaylist.SavePlaylist(.FileName, PLAYLIST_APL)
    If Right(LCase(.FileName), 3) = "m3u" Then: Call Form_Main.AdioMediaPlaylist.SavePlaylist(.FileName, PLAYLIST_M3U)
    If Right(LCase(.FileName), 3) = "pls" Then: Call Form_Main.AdioMediaPlaylist.SavePlaylist(.FileName, PLAYLIST_PLS)
    If Right(LCase(.FileName), 3) = "wpl" Then: Call Form_Main.AdioMediaPlaylist.SavePlaylist(.FileName, PLAYLIST_WPL)
End With

Exit Sub

ErrorHandler:
Select Case Err.Number
    Case 0
    Case Else: Debug.Print Err.Description
End Select
End Sub

Private Sub Button_ClearAll_Click()
Select Case FormType
    Case enumFormTypes.MidiPlayer: Form_Main.AdioMidiPlaylist.Clear
    Case enumFormTypes.Mp3Player: Form_Main.AdioMediaPlaylist.Clear
End Select
        
Call PopulatePlaylist
End Sub

Private Sub Combox_FilterType_Click()
If InitDone Then
    Dim SelectedExtension As String
    
    If Combox_FilterType.ListIndex = 0 Then
        SelectedExtension = "*"
    Else
        SelectedExtension = StrExt.SplitStr(Combox_FilterType.Text, "(*.", 1)
        SelectedExtension = Replace$(SelectedExtension, ")", vbNullString)
    End If
    
    With Form_Main
        Select Case FormType
        
            Case enumFormTypes.MidiPlayer
                If SelectedExtension <> "*" Then
                    Call .AdioMidiPlaylist.ExecQuery("extension eq '" & SelectedExtension & "'")
                Else
                    Call .AdioMidiPlaylist.ClearQuery
                End If
            
            Case enumFormTypes.Mp3Player
                If SelectedExtension <> "*" Then
                    Call .AdioMediaPlaylist.ExecQuery("extension eq '" & SelectedExtension & "'")
                Else
                    Call .AdioMediaPlaylist.ClearQuery
                End If
        End Select
    End With
    
    Call PopulatePlaylist
End If
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
Dim I As Integer
Dim SplitValue

Screen.MousePointer = vbHourglass

'Set default variables
DoneFitToSize = False
mnupopup.Visible = False
Me.Height = 8500

Call TranslateFormAndControls(Me)
Call LoadFileTypes
Call PopulatePlaylist

InitDone = True
End Sub
Private Sub Listview_Playlist_DblClick()
If Listview_Playlist.ListItems.Count > 0 Then
    Select Case FormType
        Case enumFormTypes.MidiPlayer ' Start Midiplayer
            Call Form_Main.AdioMidiPlaylist.GetTrack(PLS_GOTO, Listview_Playlist.SelectedItem)
        
        Case enumFormTypes.Mp3Player ' Start Mediaplayer
            Call Form_Main.AdioMediaPlaylist.GetTrack(PLS_GOTO, Listview_Playlist.SelectedItem)
    End Select
    
    Unload Me
End If
End Sub

Private Sub Listview_Playlist_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim I As Integer
Dim CurrentFile As String

'Add the dropped files to the playlist
For I = 1 To Data.Files.Count
    CurrentFile = Data.Files(I)
    
    Call AddToPlaylist(CurrentFile)
Next

Call FitListviewToSize
End Sub

Private Sub menu_Popup_AddDirectory_Click()
Dim I As Integer
Dim SelectedDirectory As String
Dim StringBuilder As String
Dim CurrentFile As String

'Open the browse for folder dialog
SelectedDirectory = Extensions.BrowseForFolder(Me.hwnd, GetTranslation(1060))

'Check if a folder is selected
If SelectedDirectory <> vbNullString Then
    DirToPlaylist.Path = SelectedDirectory
    
    If DirToPlaylist.ListCount > 32 Then: Call FitListviewToSize(True)
    
    'Construct the files string
    For I = 0 To DirToPlaylist.ListCount - 1
        CurrentFile = DirToPlaylist.Path & "\" & DirToPlaylist.List(I)
        
        StringBuilder = StringBuilder & CurrentFile & vbNewLine
    Next
    
    'Add the directory files to the playlist
    Call AddToPlaylist(StringBuilder)
End If
End Sub

Private Sub menu_Popup_AddFiles_Click()
Dim Files As String

On Error GoTo ErrorHandler
With CommonDialog
    .CancelError = True
    .MaxFileSize = 9999
    .DialogTitle = GetTranslation(1035)
    .Filter = ComboxboxToCommondialogFilter
    .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    .InitDir = Extensions.INIRead("main", "LastLocation", ConfigFile, App.Path)
    .ShowOpen

    If .FileName <> vbNullString Then
        Files = Extensions.CommondialogFilesToList(.FileName)
        
        Call AddToPlaylist(Files)
    End If
End With

Call Extensions.INIWrite("main", "LastLocation", Extensions.GetPathFromFilename(CommonDialog.FileName), ConfigFile)
Exit Sub

ErrorHandler:
Select Case Err.Number
    Case 0
    Case Else: Debug.Print Err.Description
End Select
End Sub

Private Sub menu_Popup_HtmlPlaylist_Click()
On Error GoTo ErrorHandler
With CommonDialog
    .DialogTitle = GetTranslation(1031)
    .CancelError = True
    .Filter = "Hyperlink Markup Langauge Page (*.html)|*.html"
    .ShowSave
    
    If .FileName <> vbNullString Then
        Call SavePlaylistAsHTML(.FileName)
    End If
End With

ErrorHandler:
Select Case Err.Number
    Case 0
    Case cdlCancel
End Select
End Sub

Private Sub menu_Popup_ShufflePlaylist_Click()
Dim I As Integer
Dim LstItem As ListItem
Dim newListItem As ListItem
    
If Listview_Playlist.ListItems.Count = 0 Then: Exit Sub
For I = 1 To Extensions.RandomNumber(1, Listview_Playlist.ListItems.Count)
    Set LstItem = Listview_Playlist.ListItems(I)
        
    Set newListItem = Listview_Playlist.ListItems.Add(, , LstItem.Text)
    newListItem.SubItems(1) = LstItem.SubItems(1)
    newListItem.SubItems(2) = LstItem.SubItems(2)
    
    Listview_Playlist.ListItems.Remove (I)
Next
End Sub

Private Sub Timer_Main_Timer()
Dim TotalSeconds As Long

If FormType = Mp3Player Then: TotalSeconds = Form_Main.AdioMediaPlaylist.GetPlaylistRuntimeInSeconds
If FormType = MidiPlayer Then: TotalSeconds = Form_Main.AdioMidiPlaylist.GetPlaylistRuntimeInSeconds

Label_PlaylistDetails.Caption = StrExt.format(GetTranslation(1026), Listview_Playlist.ListItems.Count, Extensions.SecondsToTimeSerial(TotalSeconds, LongTimeSerial))

If Listview_Playlist.ListItems.Count = 0 Then
    Button_Delete.Enabled = False
    menu_Popup_HtmlPlaylist.Enabled = False
Else
    Button_Delete.Enabled = True
    menu_Popup_HtmlPlaylist.Enabled = True
End If
End Sub
