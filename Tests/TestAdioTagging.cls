VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdioTagging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdioLibrary - Tagging Component
'
Option Explicit

Private strArtist As String
Private strTitle As String
Private strYear As String
Private strGenre As String
Private strAlbum As String
Private strDurationInSeconds As String
Private strDurationString As String
Private strBitrate As String
Private strChannels As String
Private strFileSize As String
Private strFrequency As String
Private strStereo As String

Implements ITestCase
Private Function ITestCase_TestCaseInstance() As ITestCase
    Dim objResult As New TestAdioTagging
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
    Dim objClause As New TestAdioTagging
    
    TestName = "TestAdioTagging"
    
    Call Setup
    
    Select Case TestNum
        Case objTestNum.NextNumber: Call TestDurationInSecondsPropertyRead
        Case objTestNum.NextNumber: Call TestDurationStringPropertyRead
        Case objTestNum.NextNumber: Call TestBitratePropertyRead
        Case objTestNum.NextNumber: Call TestChannelsPropertyRead
        Case objTestNum.NextNumber: Call TestFileSizePropertyRead
        Case objTestNum.NextNumber: Call TestFrequencyPropertyRead
        
        Case objTestNum.NextNumber: Call TestAlbumTagRead
        Case objTestNum.NextNumber: Call TestGenreTagRead
        Case objTestNum.NextNumber: Call TestArtistTagRead
        Case objTestNum.NextNumber: Call TestTitleTagRead
        Case objTestNum.NextNumber: Call TestYearTagRead

        Case objTestNum.NextNumber: Call TestWriteTag

        Case Else
            EndOfTests = True
    End Select
End Sub
Private Sub Setup()
With Form_Test.AdioTagging
    strArtist = .ReadTag(App.Path & "\test.mp3", tArtist, v2)
    strTitle = .ReadTag(App.Path & "\test.mp3", tTitle, v2)
    strYear = .ReadTag(App.Path & "\test.mp3", tYear, v2)
    strGenre = .ReadTag(App.Path & "\test.mp3", tGenre, v2)
    strAlbum = .ReadTag(App.Path & "\test.mp3", tAlbum, v2)
    
    strDurationInSeconds = .ReadProperty(App.Path & "\test.mp3", pDurationInSeconds)
    strDurationString = .ReadProperty(App.Path & "\test.mp3", pDurationString)
    strBitrate = .ReadProperty(App.Path & "\test.mp3", pBitrate)
    strChannels = .ReadProperty(App.Path & "\test.mp3", pChannels)
    strFileSize = .ReadProperty(App.Path & "\test.mp3", pFileSize)
    strFrequency = .ReadProperty(App.Path & "\test.mp3", pFrequency)
    strStereo = .ReadProperty(App.Path & "\test.mp3", pStereo)
End With
End Sub
Private Sub TestDurationStringPropertyRead()
AssertEqual "02:47", strDurationString, "Wrong duration amount found"
End Sub
Private Sub TestDurationInSecondsPropertyRead()
AssertEqual "167", strDurationInSeconds, "Wrong duration amount found"
End Sub
Private Sub TestAlbumTagRead()
AssertEqual "My Suno Songs", strAlbum, "Wrong album found"
End Sub
Private Sub TestGenreTagRead()
AssertEqual "Country", strGenre, "Wrong genre found"
End Sub
Private Sub TestArtistTagRead()
AssertEqual "Alex", strArtist, "Wrong artist found"
End Sub
Private Sub TestTitleTagRead()
AssertEqual "Neon Rodeo", strTitle, "Wrong title found"
End Sub
Private Sub TestYearTagRead()
AssertEqual "2026", strYear, "Wrong year found"
End Sub
Private Sub TestBitratePropertyRead()
AssertEqual "190", strBitrate, "Bitrate should be 190"
End Sub
Private Sub TestChannelsPropertyRead()
AssertEqual "Stereo", strChannels, "Channels should be the value `Stereo`"
End Sub
Private Sub TestFileSizePropertyRead()
AssertEqual "3992008", strFileSize, "File size should be 3992008 bytes"
End Sub
Private Sub TestFrequencyPropertyRead()
AssertEqual "48000", strFrequency, "Frequency should be 48000"
End Sub
Private Sub TestWriteTag()
With Form_Test.AdioTagging
    AssertTrue .WriteTag(App.Path & "\test.mp3", "Test Tag", tComposer, v2), "Writing of test tag must succeed"
    AssertEqual "Test Tag", .ReadTag(App.Path & "\test.mp3", tComposer, v2), "Tag should be `Test Tag` after writing"
    
    AssertTrue .WriteTag(App.Path & "\test.mp3", vbNullString, tComposer, v2), "Writing of empty tag must succeed"
    AssertEqual vbNullString, .ReadTag(App.Path & "\test.mp3", tComposer, v2), "Tag should be empty"
End With
End Sub
