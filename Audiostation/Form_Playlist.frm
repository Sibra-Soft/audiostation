VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1015"
   Begin Audiostation.ButtonBig cmdSave 
      Height          =   390
      Left            =   10200
      TabIndex        =   7
      Top             =   6720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
      Caption         =   "Save"
   End
   Begin MSComctlLib.ListView lstPlaylist 
      Height          =   6735
      Left            =   120
      TabIndex        =   18
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
      Left            =   10200
      TabIndex        =   17
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Trm_Count 
      Interval        =   200
      Left            =   10680
      Top             =   3480
   End
   Begin VB.ComboBox cboFileTypes 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   8655
   End
   Begin Audiostation.ButtonBig cmdClose 
      Height          =   390
      Left            =   10200
      TabIndex        =   6
      Tag             =   "1001"
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
      Caption         =   "Close"
   End
   Begin VB.Timer Trm_Enable 
      Interval        =   1
      Left            =   10200
      Top             =   3480
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   10200
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
   Begin VB.PictureBox ButtonBar 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   0
      Left            =   10080
      ScaleHeight     =   2535
      ScaleWidth      =   1935
      TabIndex        =   10
      Top             =   840
      Width           =   1935
      Begin Audiostation.ButtonBig cmdDelete 
         Height          =   390
         Left            =   120
         TabIndex        =   11
         Tag             =   "1004"
         Top             =   1440
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
      End
      Begin Audiostation.ButtonBig cmdClear 
         Height          =   390
         Left            =   120
         TabIndex        =   12
         Tag             =   "1005"
         Top             =   0
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         Caption         =   "Clear All"
      End
      Begin Audiostation.ButtonBig isButton6 
         Height          =   390
         Left            =   120
         TabIndex        =   13
         Tag             =   "1009"
         Top             =   960
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
      End
      Begin Audiostation.ButtonBig isButton5 
         Height          =   390
         Left            =   120
         TabIndex        =   14
         Tag             =   "1008"
         Top             =   480
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
      End
      Begin Audiostation.ButtonBig cmdPlaylistOptions 
         Height          =   390
         Left            =   120
         TabIndex        =   15
         Tag             =   "1054"
         Top             =   2040
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
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
      TabIndex        =   19
      Top             =   600
      Width           =   9975
   End
   Begin VB.Label lblPlaylistCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 File(s) in this playlist, total duration 00:00:00"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   7740
      Width           =   4020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Playlist:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   1155
   End
   Begin VB.Menu mnupopup 
      Caption         =   "- POPUP -"
      Begin VB.Menu mnuaddfile_popup 
         Caption         =   "&Add File(s)"
      End
      Begin VB.Menu mnuadddirectory_popup 
         Caption         =   "&Add Directory"
      End
      Begin VB.Menu space01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhtmlplaylist_popup 
         Caption         =   "&Generate HTML playlist"
      End
   End
End
Attribute VB_Name = "Form_Playlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const LB_FINDSTRING = &H18F

Dim DoneFitToSize As Boolean
Dim FileTypesToShow As String
Dim strLocation As String

Enum PlsType
    [Windows Media Player]
    [ShoutCast Playlist]
    [Common Playlist]
    [Audiostation Playlist]
End Enum

Public CurrentFormType As enumFormTypes
Private Sub FitListviewToSize(Optional Force As Boolean)
'Do nothing if it's already been done
If DoneFitToSize = True Then: Exit Sub

'Check if the scrollbar is visible
If lstPlaylist.ListItems.Count > 32 Or Force = True Then
    Form_Playlist.lstPlaylist.ColumnHeaders(1).Width = Form_Playlist.lstPlaylist.ColumnHeaders(1).Width - 250
    DoneFitToSize = True
End If
End Sub
Private Function ComboxboxToCommondialogFilter() As String
Dim I As Integer
Dim FileType As String
Dim Description As String
Dim FilterString As New StringBuilder

For I = 0 To cboFileTypes.ListCount - 1
    FileType = Extensions.StringBetween("(", ")", cboFileTypes.List(I))
    Description = Extensions.Explode(cboFileTypes.List(I), "(", 0)
    
    FilterString.Append Description & "(" & FileType & ")|" & FileType & "|"
Next

ComboxboxToCommondialogFilter = FilterString.toString
End Function
Private Sub SavePlaylistAsHTML(FileName As String)
Dim I As Integer
Dim TemplateContent As String
Dim table As New StringBuilder

'Open the file with the template content
TemplateContent = Extensions.FileGetContents(App.path & "\templates\html_playlist.tpl")

'Create the table
table.AppendNL "<table style='width:100%;' border='2px' >"
table.AppendNL "<thead>"
    table.AppendNL "<tr>"
        table.AppendNL "<td><strong>Nr.</strong></td>"
        table.AppendNL "<td><strong>Filename</strong></td>"
        table.AppendNL "<td><strong>Duration</strong></td>"
    table.AppendNL "</tr>"
table.AppendNL "</thead>"

For I = 1 To lstPlaylist.ListItems.Count
    table.AppendNL "<tr>"
        table.AppendNL "<td>" & lstPlaylist.ListItems(I).Text & "</td>"
        table.AppendNL "<td>" & lstPlaylist.ListItems(I).SubItems(1) & "</td>"
        table.AppendNL "<td>" & lstPlaylist.ListItems(I).SubItems(2) & "</td>"
    table.AppendNL "</tr>"
Next

table.AppendNL "</table>"

'Replace the template variable with the created table
TemplateContent = Replace(TemplateContent, "{PlaylistTable}", table.toString)

Call Extensions.FilePutContents(FileName, TemplateContent)
End Sub
Public Sub AddToPlaylist(file As String)
Dim MediaDuration As String
Dim lstItem As ListItem
Dim Files
Dim FilesToAddCount As Integer
Dim I As Integer
Dim FirstIndex As Integer
Dim AddingMoreThanOne As Boolean
Dim FileNotFoundCount As Integer
Dim MediaTagManager As New Mp3Info

On Error GoTo ErrorHandler
Files = Split(file, vbNewLine)
FilesToAddCount = UBound(Files)

If FilesToAddCount > 0 Then: FirstIndex = 0
If FilesToAddCount > 1 Then: AddingMoreThanOne = True

'More than one file, so load the busy form
If AddingMoreThanOne = True Then
    Screen.MousePointer = vbHourglass
    
    Form_Busy.ProgressBar.Value = 0
    Form_Busy.ProgressBar.Max = FilesToAddCount
    Form_Busy.Show
End If

'Add files to the playlist
For I = FirstIndex To FilesToAddCount
    'Get current file to process
    file = Files(I)
    
    If Dir(file, vbDirectory) = vbNullString Then
        Debug.Print "File not found"
        FileNotFoundCount = FileNotFoundCount + 1
    Else
        If Not file = vbNullString Then
            MediaDuration = 0
        
            If CurrentFormType = Mp3Player Then
                If (LCase(Right(file, 3)) = "mp3") Then
                    MediaTagManager.FileName = file
                    MediaDuration = Extensions.TimeString(MediaTagManager.SongLength)
                End If
            End If
            
            If MediaDuration = "0" Then: MediaDuration = "-"
            
            'Add the file to the listview
            Set lstItem = lstPlaylist.ListItems.Add(, file, Format(lstPlaylist.ListItems.Count + 1, "00"))
                lstItem.SubItems(1) = file
                lstItem.SubItems(2) = MediaDuration
        End If
        
        If AddingMoreThanOne = True Then
            DoEvents
            Form_Busy.ProgressBar.Value = Form_Busy.ProgressBar.Value + 1
        End If
    End If
Next

'Check if all files are found
If FileNotFoundCount > 0 Then MsgBox Extensions.StringFormat(GetLanguage(1068), FileNotFoundCount), vbExclamation

'Restore default
Screen.MousePointer = vbDefault
Unload Form_Busy

Exit Sub
ErrorHandler:
Select Case err.Number
    Case 0
    Case 35602
    Case Else: Debug.Print err.Number & " - " & err.Description
End Select
End Sub
Private Function SavePlaylist(ByVal strFile As String, ByRef lstFormList As ListView, ByVal PlaylistType As PlsType, Optional ByVal blnClearList As Boolean = False)
Dim I As Integer
Dim FN As Integer

Dim MusicLibraryLocation As String
Dim strFilename As String
Dim lstItem As ListItem

strLocation = strFile
strFilename = Extensions.GetFileNameFromFilePath(strFile, False)

FN = FreeFile

Select Case PlaylistType
    Case 0
        Dim PlaylistName As String
        
        PlaylistName = Extensions.GetFileNameFromFile(strFile)
        
        Open strFile For Output As #1
          Print #1, "<?wpl version="; 1#; "?>"
          Print #1, "<smil>"
          Print #1, "    <head>"
          Print #1, "        <title>" & PlaylistName & "</title>"
          Print #1, "    </head>"
          Print #1, "    <body>"
          Print #1, "        <seq>"
          
          ' Get all the items from the selected playlist
          For Each lstItem In lstFormList.ListItems
            Print #1, "<media src=""" & lstItem.SubItems(1) & """/>"
          Next

          Print #1, "       </seq>"
          Print #1, "    </body>"
          Print #1, "</smil>"
        Close #1
        
    Case 1
        For I = 1 To lstFormList.ListItems.Count
            Call Extensions.INIWrite("playlist", "File" & I, lstFormList.ListItems(I).Key, strFile)
        Next I
        
        Call Extensions.INIWrite("playlist", "NumberOfEntries", lstFormList.ListItems.Count, strFile)
        Call Extensions.INIWrite("playlist", "Version", 2, strFile)
    
    Case 2
        Open strFile For Output As #FN
            Print #FN, "#EXTM3U"
            
            For I = 1 To lstFormList.ListItems.Count
              Print #FN, "#EXTINF:0, " & Extensions.GetFileNameFromFilePath(lstFormList.ListItems(I).Key, False)
              Print #FN, lstFormList.ListItems(I).Key
              Print #FN, ""
            Next I
        Close #FN
    
    Case 3
        Open strFile For Output As #FN
            For I = 1 To lstFormList.ListItems.Count
              Print #FN, lstFormList.ListItems(I).Key
            Next I
        Close #FN
        
        If blnClearList = True Then lstFormList.ListItems.Clear
End Select
End Function
Public Function PopulatePlaylist(Playlist As LocalStorage)
Dim fso As New FileSystemObject
Dim I As Integer

Dim CurrentFile As String

lstPlaylist.ListItems.Clear

'Get the playlist from the local saved storage
For I = 1 To Playlist.StorageContainer.Count
    CurrentFile = Playlist.GetItemByIndex(I, 1)
    
    'Add the file to the listview
    Set lstItem = lstPlaylist.ListItems.Add(, CurrentFile, Playlist.GetItemByIndex(I, 0))
        lstItem.SubItems(1) = CurrentFile
        lstItem.SubItems(2) = Playlist.GetItemByIndex(I, 2)
    
    'Select the current playing mp3 audio track
    If AudiostationMP3Player.CurrentTrackNumber = I And CurrentFormType = Mp3Player Then
        lstPlaylist.ListItems(I).Bold = True
        lstPlaylist.ListItems(I).ListSubItems(1).Bold = True
        lstPlaylist.ListItems(I).ListSubItems(2).Bold = True
    End If
    
    'Select the current playing midi audio track
    If AudiostationMidiPlayer.CurrentMidiTrackNumber = I And CurrentFormType = MidiPlayer Then
        lstPlaylist.ListItems(I).Bold = True
        lstPlaylist.ListItems(I).ListSubItems(1).Bold = True
        lstPlaylist.ListItems(I).ListSubItems(2).Bold = True
    End If
Next

Call FitListviewToSize
End Function
Public Function LoadFileTypes()
Select Case CurrentFormType
    Case enumFormTypes.MidiPlayer
        cboFileTypes.AddItem GetLanguage(1044) & " (*.mid;*.midi;*.kar;*.mus;*.sid)"
        cboFileTypes.AddItem "MIDI File (*.mid)"
        cboFileTypes.AddItem "MIDI File (*.midi)"
        cboFileTypes.AddItem "Karaoke File (*.kar)"
        cboFileTypes.AddItem "Sibra-Soft Beep Symphony (*.mus)"
        cboFileTypes.AddItem "Commodore64 Sound File (*.sid)"
        cboFileTypes.AddItem GetLanguage(1016) & " (*.*)"
        
        cboFileTypes.ListIndex = 0 'Select the first item
        
    Case enumFormTypes.Mp3Player
        cboFileTypes.AddItem GetLanguage(1044) & " (*.mp3;*.mp2;*.m4a;*.ra;*.cda;*.wav;*.aif;*.aac;*.snd,;*.au;*.wma;*.rmi)"
        cboFileTypes.AddItem "MPEG-1 Layer 3 (*.mp3)"
        cboFileTypes.AddItem "MPEG-1 Layer 2 (*.mp2)"
        cboFileTypes.AddItem "MPEG-4 Layer 4 Audio (*.m4a)"
        cboFileTypes.AddItem "Convert - Real Audio (*.ra)"
        cboFileTypes.AddItem "Convert - Real Media (*.rm)"
        cboFileTypes.AddItem "CD Audio (*.cda)"
        cboFileTypes.AddItem "Microsoft WaveForm Audio (*.wav)"
        cboFileTypes.AddItem "Audio Interchange File (*.aif)"
        cboFileTypes.AddItem "Advanced Audio Coding (*.aac)"
        cboFileTypes.AddItem "Sun Microsystems Sound (*.snd)"
        cboFileTypes.AddItem "Sun Microsystems Audio (*.au)"
        cboFileTypes.AddItem "Windows Media Audio (*.wma)"
        cboFileTypes.AddItem "Musical Instrument Digital Interface (*.rmi)"
        cboFileTypes.AddItem "Convert - Voice File Format (*.act)"
        cboFileTypes.AddItem "Convert - Apple Core Format (*.caf)"
        cboFileTypes.AddItem "Convert - Westwood Studios Audio (*.wsaud)"
        cboFileTypes.AddItem "Convert - Sony Wave64 (*.w64)"
        cboFileTypes.AddItem "Convert - OGG (*.ogg)"
        cboFileTypes.AddItem "Convert - Sony Opend Audio (*.amo)"
        cboFileTypes.AddItem "Convert - Creative Voice (*.voc)"
        cboFileTypes.AddItem GetLanguage(1016) & " (*.*)"
        
        cboFileTypes.ListIndex = 0 'Select the first item
End Select
End Function

Private Sub cboFileTypes_Click()
Dim SplitValue
Dim FileTypeArray As String
Dim CurrentPlaylist As LocalStorage

Select Case CurrentFormType
    Case enumFormTypes.MidiPlayer: Set CurrentPlaylist = MidiPlaylist
    Case enumFormTypes.Mp3Player: Set CurrentPlaylist = Mp3Playlist
End Select

'First save the current playlist to the storage
Call CurrentPlaylist.ListviewToStorage(lstPlaylist, 1)

'Get the current file type to filter
SplitValue = Split(cboFileTypes.Text, "(*")
FileTypeArray = Replace(SplitValue(1), ")", vbNullString)

If Not cboFileTypes.ListIndex = 0 Then: CurrentPlaylist.Filter (FileTypeArray)

'Rebuild the playlist
Call PopulatePlaylist(CurrentPlaylist)

'Check if the filter is set, yes then clear the current filter
If CurrentPlaylist.IsFilterd = True Then: CurrentPlaylist.ClearFilter
End Sub

Private Sub cmdClear_Click()
lstPlaylist.ListItems.Clear
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
lstPlaylist.ListItems.Remove (lstPlaylist.SelectedItem.Index)
End Sub

Private Sub cmdPlaylistOptions_Click()
PopupMenu mnupopup
End Sub

Private Sub cmdSave_Click()
Select Case CurrentFormType
    Case enumFormTypes.MidiPlayer
        Call MidiPlaylist.ClearStorage
        Call MidiPlaylist.ListviewToStorage(lstPlaylist, 1)
        
    Case enumFormTypes.Mp3Player
        Call Mp3Playlist.ClearStorage
        Call Mp3Playlist.ListviewToStorage(lstPlaylist, 1)
End Select

Unload Me
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

Call SetLanguage(Me)
Call LoadFileTypes

Select Case CurrentFormType
    Case enumFormTypes.MidiPlayer: Call PopulatePlaylist(MidiPlaylist)
    Case enumFormTypes.Mp3Player: Call PopulatePlaylist(Mp3Playlist)
End Select
End Sub

Private Sub isButton5_Click()
On Error GoTo ErrorHandler
With CommonDialog
    .CancelError = True
    .FileName = App.path
    .DialogTitle = GetLanguage(1017)
    .Filter = "Audiostation Playlist (*.apl)|*.apl|" & GetLanguage(1019) & " (.m3u)|*.m3u|ShoutCast Playlist (*.pls)|*.pls|Windows Media Player Playlist (*.wpl)|*.wpl"
    .ShowSave

    If Right(LCase(.FileName), 3) = "apl" Then: Call SavePlaylist(.FileName, lstPlaylist, [Audiostation Playlist])
    If Right(LCase(.FileName), 3) = "m3u" Then: Call SavePlaylist(.FileName, lstPlaylist, [Common Playlist])
    If Right(LCase(.FileName), 3) = "pls" Then: Call SavePlaylist(.FileName, lstPlaylist, [ShoutCast Playlist])
    If Right(LCase(.FileName), 3) = "wpl" Then: Call SavePlaylist(.FileName, lstPlaylist, [Windows Media Player])
End With

Exit Sub

ErrorHandler:
Select Case err.Number
    Case 0
    Case Else: Debug.Print err.Description
End Select
End Sub

Private Sub isButton6_Click()
On Error GoTo ErrorHandler
With CommonDialog
    .CancelError = True
    .FileName = App.path
    .DialogTitle = GetLanguage(1018)
    .Filter = "Audiostation Playlist (*.apl)|*.apl|" & GetLanguage(1019) & " (.m3u)|*.m3u|ShoutCast Playlist (*.pls)|*.pls|Windows Media Player Playlist (*.wpl)|*.wpl"
    .ShowOpen
    
    If .FilterIndex = 1 Then: Call ModPlaylist.OpenAplPlaylist(.FileName)
    If .FilterIndex = 2 Then: Call ModPlaylist.OpenM3uPlaylist(.FileName)
    If .FilterIndex = 3 Then: Call ModPlaylist.OpenPlsPlaylist(.FileName)
    If .FilterIndex = 4 Then: Call ModPlaylist.OpenWplPlaylist(.FileName)
    
    Call FitListviewToSize
End With

Exit Sub

ErrorHandler:
Select Case err.Number
    Case 0
    Case Else: Debug.Print err.Description
End Select
End Sub

Private Sub lstPlaylist_DblClick()
If lstPlaylist.ListItems.Count > 0 Then
    Call Mp3Playlist.ListviewToStorage(lstPlaylist, 1)
    
    Select Case CurrentFormType
        Case enumFormTypes.MidiPlayer
            AudiostationMidiPlayer.CurrentMidiTrackNumber = lstPlaylist.SelectedItem.Index
            AudiostationMidiPlayer.StartMidiPlayback
            
        Case enumFormTypes.Mp3Player
            AudiostationMP3Player.CurrentTrackNumber = lstPlaylist.SelectedItem.Index
            AudiostationMP3Player.StartPlay
    End Select
    
    Unload Me
End If
End Sub

Private Sub lstPlaylist_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
Dim CurrentFile As String

'Add the dropped files to the playlist
For I = 1 To Data.Files.Count
    CurrentFile = Data.Files(I)
    
    Call AddToPlaylist(CurrentFile)
Next

Call FitListviewToSize
End Sub

Private Sub mnuadddirectory_popup_Click()
Dim I As Integer
Dim SelectedDirectory As String
Dim StringBuilder As String
Dim CurrentFile As String

'Open the browse for folder dialog
SelectedDirectory = Extensions.BrowseForFolder(Me.hwnd, "Open folder")

'Check if a folder is selected
If SelectedDirectory <> vbNullString Then
    DirToPlaylist.path = SelectedDirectory
    
    If DirToPlaylist.ListCount > 32 Then: Call FitListviewToSize(True)
    
    'Construct the files string
    For I = 0 To DirToPlaylist.ListCount - 1
        CurrentFile = DirToPlaylist.path & "\" & DirToPlaylist.List(I)
        
        StringBuilder = StringBuilder & CurrentFile & vbNewLine
    Next
    
    'Add the directory files to the playlist
    Call AddToPlaylist(StringBuilder)
End If
End Sub

Private Sub mnuaddfile_popup_Click()
Dim Files As String

On Error GoTo ErrorHandler
With CommonDialog
    .CancelError = True
    .MaxFileSize = 9999
    .DialogTitle = "Add file(s) to playlist"
    .Filter = ComboxboxToCommondialogFilter
    .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
    .ShowOpen

    If .FileName <> vbNullString Then
        Files = Extensions.CommondialogFilesToList(.FileName)
        
        Call AddToPlaylist(Files)
    End If
End With

ErrorHandler:
Select Case err.Number
    Case 0
    Case Else: Debug.Print err.Description
End Select
End Sub

Private Sub mnuhtmlplaylist_popup_Click()
On Error GoTo ErrorHandler
With CommonDialog
    .DialogTitle = "Generate HTML playlist"
    .CancelError = True
    .Filter = "Hyperlink Markup Langauge Page (*.html)|*.html"
    .ShowSave
    
    If .FileName <> vbNullString Then
        Call SavePlaylistAsHTML(.FileName)
    End If
End With

ErrorHandler:
Select Case err.Number
    Case 0
    Case cdlCancel
End Select
End Sub
Private Sub Trm_Count_Timer()
Dim Minutes As Long
Dim seconds As Long
Dim SplitValue
Dim I As Integer
Dim TotalSeconds As Long
Dim ItemValue As String

For I = 1 To lstPlaylist.ListItems.Count
    ItemValue = lstPlaylist.ListItems(I).SubItems(2)
    
    If ItemValue = "-" Or ItemValue = vbNullString Then
        'Do nothing
    Else
        SplitValue = Split(ItemValue, ":")
        
        Minutes = Minutes + SplitValue(0)
        seconds = seconds + SplitValue(1)
    End If
Next

TotalSeconds = Minutes * 60 + seconds

lblPlaylistCount.Caption = Extensions.StringFormat(GetLanguage(1053), lstPlaylist.ListItems.Count, Extensions.SecondsToTimeSerial(TotalSeconds, LongTimeSerial))
End Sub

Private Sub Trm_Enable_Timer()
If lstPlaylist.ListItems.Count = 0 Then
    cmdDelete.Enabled = False
    mnuhtmlplaylist_popup.Enabled = False
Else
    cmdDelete.Enabled = True
    mnuhtmlplaylist_popup.Enabled = True
End If
End Sub
