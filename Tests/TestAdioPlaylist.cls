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
        
        Case objTestNum.NextNumber: Call TestSavePlaylist
        Case objTestNum.NextNumber: Call TestLoadPlaylistFile
        Case objTestNum.NextNumber: Call TestPlaylistQuery
        Case objTestNum.NextNumber: Call TestTotalPlaylistDuration
        
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
Form_Test.AdioPlaylist.AddFile App.Path & "\test.mp3"
AssertEqual Form_Test.AdioPlaylist.ListCount, 1, "Playlist must contain 1 item"
End Sub
Private Sub TestRemoveItemFromPlaylist()
Form_Test.AdioPlaylist.AddFile App.Path & "\test.mp3"
AssertEqual Form_Test.AdioPlaylist.ListCount, 1, "Playlist must contain 1 item"
Form_Test.AdioPlaylist.Remove 1
AssertEqual Form_Test.AdioPlaylist.ListCount, 0, "Playlist should be empty"
End Sub
Private Sub TestGetFirstPlaylistTrack()
Form_Test.AdioPlaylist.AddFile App.Path & "\test.mp3"
AssertEqual App.Path & "\test.mp3", Form_Test.AdioPlaylist.GetTrack(PLS_FIRST).LocalFile, "Wrong file for first item in the playlist"
End Sub
Private Sub TestReadPlaylistItemProperties()
Dim plsItem As mdlAdioPlaylistItem

Form_Test.AdioPlaylist.AddFile App.Path & "\test.mp3"

With Form_Test.AdioPlaylist
    Set plsItem = .GetTrack(PLS_FIRST)
    
    AssertEqual "mp3", plsItem.FileExtension, "Playlist item extension property should be `mp3`"
    AssertEqual "test.mp3", plsItem.FileName, "Playlist item filename property should be `test.mp3`"
    AssertEqual 1, plsItem.nR, "Playlist item Nr property should be 0"
End With
End Sub
Private Sub TestSavePlaylist()
With Form_Test.AdioPlaylist
    .AddFile App.Path & "\test.mp3"
    .AddFile App.Path & "\test2.mp3"
    .AddFile App.Path & "\test3.mp3"
    
    AssertEqual 3, .ListCount, "Playlist must contain 3 tracks"

    AssertTrue .SavePlaylist(App.Path & "\test.apl", PLAYLIST_APL), "A playlist file should be saved"
End With
End Sub
Private Sub TestLoadPlaylistFile()
With Form_Test.AdioPlaylist
    .LoadPlaylist App.Path & "\test.apl", PLAYLIST_APL
    
    AssertEqual 3, .ListCount, "Playlist should contain 3 tracks"
End With
End Sub
Private Sub TestPlaylistQuery()
With Form_Test.AdioPlaylist
    .LoadPlaylist App.Path & "\test.apl", PLAYLIST_APL
    
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
    .LoadPlaylist App.Path & "\test.apl", PLAYLIST_APL
    
    AssertEqual 512, .GetPlaylistRuntimeInSeconds, "Wrong total length of the playlist found"
    AssertEqual "00:08:32", .GetPlaylistRuntimeString, "Wrong duration string for total length of the playlist"
End With
End Sub
