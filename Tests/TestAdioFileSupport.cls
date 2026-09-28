VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdioFileSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdioLibrary - File Support Test
'
Option Explicit

Implements ITestCase
Private Function ITestCase_TestCaseInstance() As ITestCase
    Dim objResult As New TestAdioFileSupport
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
    Dim objClause As New TestAdioFileSupport
    
    TestName = "TestAdioFileSupport"
    
    Select Case TestNum
        Case tstcTestActionSetup: Call Setup
        
        Case objTestNum.NextNumber: Call TestUnsupportedFileError
        Case objTestNum.NextNumber: Call TestFileNotFoundError
        
        Case objTestNum.NextNumber: Call TestMP3FileSupport
        Case objTestNum.NextNumber: Call TestWAVFileSupport
        Case objTestNum.NextNumber: Call TestFLACFileSupport
        Case objTestNum.NextNumber: Call TestOGGFileSupport
        
        Case Else
            EndOfTests = True
    End Select
End Sub
Private Sub Setup()
Form_Test.MediaPlayerErrorCount = 0
End Sub
Private Sub TestFileNotFoundError()
With Form_Test
    .AdioMediaPlayer.LoadFile App.Path & "\not_found.mp3"

    AssertTrue (.MediaPlayerErrorCount > 0), "Error should be thrown when file could not be found"
End With
End Sub
Private Sub TestUnsupportedFileError()
With Form_Test
    .AdioMediaPlayer.LoadFile App.Path & "\bass.dll"
    
    AssertTrue (.MediaPlayerErrorCount > 0), "Error should be thrown when unsupported file want to load"
End With
End Sub
Private Sub TestMP3FileSupport()
With Form_Test
    .AdioMediaPlayer.LoadFile App.Path & "\test.mp3"

    AssertEqual 0, .MediaPlayerErrorCount, "No errors should be thrown when loading a MP3 file"
End With
End Sub
Private Sub TestWAVFileSupport()
With Form_Test
    .AdioMediaPlayer.LoadFile App.Path & "\test.wav"

    AssertEqual 0, .MediaPlayerErrorCount, "No errors should be thrown when loading a WAV file"
End With
End Sub
Private Sub TestFLACFileSupport()
With Form_Test
    .AdioMediaPlayer.LoadFile App.Path & "\test.flac"

    AssertEqual 0, .MediaPlayerErrorCount, "No errors should be thrown when loading a FLAC file"
End With
End Sub
Private Sub TestOGGFileSupport()
With Form_Test
    .AdioMediaPlayer.LoadFile App.Path & "\test.ogg"

    AssertEqual 0, .MediaPlayerErrorCount, "No errors should be thrown when loading a OGG file"
End With
End Sub
