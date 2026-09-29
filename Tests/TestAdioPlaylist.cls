VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdioPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdioLibrary - Playlist Component
'
Option Explicit

Implements ITestCase

Private mblnSetupRan As Boolean
Private Function ITestCase_TestCaseInstance() As ITestCase
    Dim objResult As New TestAdioPlaylist
    Debug.Assert TypeName(objResult) = TypeName(Me)
    Set ITestCase_TestCaseInstance = objResult
End Function

Private Property Get ITestCase_Suite() As Variant
    ' Do nothing here
    ' Only run tests within RunTest.
End Property
Private Sub ITestCase_RunTest( _
    ByVal TestNum As Long, _
    TestName As String, _
    ExpectErrNum As Long, _
    EndOfTests As Boolean _
)
    Dim objTestNum As New TestNumGenerator
    Dim objClause As New TestAdioPlaylist
    
    TestName = "TestAdioPlaylist"
    
    Select Case TestNum
        Case tstcTestActionSetup: Call Setup
        
        Case objTestNum.NextNumber: Call TestAddItemToPlaylist
        Case objTestNum.NextNumber: Call TestRemoveItemFromPlaylist
        Case objTestNum.NextNumber: Call TestGetFirstPlaylistTrack
        Case objTestNum.NextNumber: Call TestReadPlaylistItemProperties
        
        Case objTestNum.NextNumber: Call TestSaveAPLPlaylist                ' Audiostatio Playlist
        Case objTestNum.NextNumber: Call TestLoadAPLPlaylistFile
        
        Case objTestNum.NextNumber: Call TestPlaylistQuery
        Case objTestNum.NextNumber: Call TestTotalPlaylistDuration
        
        'Case objTestNum.NextNumber: Call TestSaveWPLPlaylist                ' Windows Media Player Playlist
        'Case objTestNum.NextNumber: Call TestSaveM3UPlaylist                ' M3U Playlist
        'Case objTestNum.NextNumber: Call TestSavePLSPlaylist                ' PLS Playlist
        
        'Case objTestNum.NextNumber: Call TestLoadWPLPlaylistFile
        'Case objTestNum.NextNumber: Call TestLoadM3UPlaylistFile
        'Case objTestNum.NextNumber: Call TestLoadPLSPlaylistFile
        
        Case Else
            EndOfTests = True
    End Select
End Sub
Private Sub Setup()
mblnSetupRan = True

Form_Test.AdioPlaylist.Clear
End Sub
Private Sub TestAddItemToPlaylist()
AssertEqual Form_Test.AdioPlaylist.ListCount, 0, "Playlist must be empty before adding"
Form_Test.AdioPlaylist.AddFile strTestFilesPath & "\test.mp3"
AssertEqual Form_Test.AdioPlaylist.ListCount, 1, "Playlist must contain 1 item"
End Sub
Private Sub TestRemoveItemFromPlaylist()
Form_Test.AdioPlaylist.AddFile strTestFilesPath & "\test.mp3"
AssertEqual Form_Test.AdioPlaylist.ListCount, 1, "Playlist must contain 1 item"
Form_Test.AdioPlaylist.Remove 1
AssertEqual Form_Test.AdioPlaylist.ListCount, 0, "Playlist should be empty"
End Sub
Private Sub TestGetFirstPlaylistTrack()
Form_Test.AdioPlaylist.AddFile strTestFilesPath & "\test.mp3"
AssertEqual strTestFilesPath & "\test.mp3", Form_Test.AdioPlaylist.GetTrack(PLS_FIRST).plsLocalFile, "Wrong file for first item in the playlist"
End Sub
Private Sub TestReadPlaylistItemProperties()
Dim plsItem As mdlAdioPlaylistItem

Form_Test.AdioPlaylist.AddFile strTestFilesPath & "\test.mp3"

With Form_Test.AdioPlaylist
    Set plsItem = .GetTrack(PLS_FIRST)
    
    AssertEqual "mp3", plsItem.plsFileExtension, "Playlist item extension property should be `mp3`"
    AssertEqual "test.mp3", plsItem.plsFileName, "Playlist item filename property should be `test.mp3`"
    AssertEqual 1, plsItem.plsNr, "Playlist item Nr property should be 0"
End With
End Sub
Private Sub TestSaveWPLPlaylist()
With Form_Test.AdioPlaylist
    .AddFile strTestFilesPath & "\test.mp3"
    .AddFile strTestFilesPath & "\test2.mp3"
    .AddFile strTestFilesPath & "\test3.mp3"
    
    AssertEqual 3, .ListCount, "Playlist must contain 3 tracks"

    AssertTrue .SavePlaylist(strTestFilesPath & "\test.wpl", PLAYLIST_WPL), "A playlist file should be saved"
End With
End Sub
Private Sub TestSaveAPLPlaylist()
With Form_Test.AdioPlaylist
    .AddFile strTestFilesPath & "\test.mp3"
    .AddFile strTestFilesPath & "\test2.mp3"
    .AddFile strTestFilesPath & "\test3.mp3"
    
    AssertEqual 3, .ListCount, "Playlist must contain 3 tracks"

    AssertTrue .SavePlaylist(strTestFilesPath & "\test.apl", PLAYLIST_APL), "A playlist file should be saved"
End With
End Sub
Private Sub TestSavePLSPlaylist()
With Form_Test.AdioPlaylist
    .AddFile strTestFilesPath & "\test.mp3"
    .AddFile strTestFilesPath & "\test2.mp3"
    .AddFile strTestFilesPath & "\test3.mp3"
    
    AssertEqual 3, .ListCount, "Playlist must contain 3 tracks"

    AssertTrue .SavePlaylist(strTestFilesPath & "\test.pls", PLAYLIST_PLS), "A playlist file should be saved"
End With
End Sub
Private Sub TestSaveM3UPlaylist()
With Form_Test.AdioPlaylist
    .AddFile strTestFilesPath & "\test.mp3"
    .AddFile strTestFilesPath & "\test2.mp3"
    .AddFile strTestFilesPath & "\test3.mp3"
    
    AssertEqual 3, .ListCount, "Playlist must contain 3 tracks"

    AssertTrue .SavePlaylist(strTestFilesPath & "\test.m3u", PLAYLIST_M3U), "A playlist file should be saved"
End With
End Sub
Private Sub TestLoadAPLPlaylistFile()
With Form_Test.AdioPlaylist
    .LoadPlaylist strTestFilesPath & "\test.apl", PLAYLIST_APL
    
    AssertEqual 3, .ListCount, "APL Playlist should contain 3 tracks"
End With
End Sub
Private Sub TestLoadWPLPlaylistFile()
With Form_Test.AdioPlaylist
    .LoadPlaylist strTestFilesPath & "\test.wpl", PLAYLIST_WPL
    
    AssertEqual 3, .ListCount, "WPL Playlist should contain 3 tracks"
End With
End Sub
Private Sub TestLoadM3UPlaylistFile()
With Form_Test.AdioPlaylist
    .LoadPlaylist strTestFilesPath & "\test.m3u", PLAYLIST_M3U
    
    AssertEqual 3, .ListCount, "M3U Playlist should contain 3 tracks"
End With
End Sub
Private Sub TestLoadPLSPlaylistFile()
With Form_Test.AdioPlaylist
    .LoadPlaylist strTestFilesPath & "\test.pls", PLAYLIST_PLS
    
    AssertEqual 3, .ListCount, "PLS Playlist should contain 3 tracks"
End With
End Sub
Private Sub TestPlaylistQuery()
With Form_Test.AdioPlaylist
    .LoadPlaylist strTestFilesPath & "\test.apl", PLAYLIST_APL
    
    .ClearQuery
    .ExecQuery "extension test 'mp3'"
    
    AssertEqual 1, Form_Test.PlaylistErrorCount, "Last query should return error"
    
    .ClearQuery
    .ExecQuery "extension eq 'wav'"
    
    AssertTrue .QueryActive, "A query should be active at this point"
    AssertEqual 0, .GetList.Count, "The specified query should return 0 results"
    
    .ClearQuery
    .ExecQuery "extension like 'mp3'"
    
    AssertTrue .QueryActive, "A query should be active at this point"
    AssertEqual 3, .GetList.Count, "The specified query should return 3 results"
End With
End Sub
Private Sub TestTotalPlaylistDuration()
With Form_Test.AdioPlaylist
    .LoadPlaylist strTestFilesPath & "\test.apl", PLAYLIST_APL
    
    AssertEqual 512, .GetPlaylistRuntimeInSeconds, "Wrong total length of the playlist found"
    AssertEqual "00:08:32", .GetPlaylistRuntimeString, "Wrong duration string for total length of the playlist"
End With
End Sub
