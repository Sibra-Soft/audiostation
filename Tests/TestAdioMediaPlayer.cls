VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdioMediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdioLibrary - MediaPlayer Component
'
Option Explicit

Implements ITestCase
Private Function ITestCase_TestCaseInstance() As ITestCase
    Dim objResult As New TestAdioMediaPlayer
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
    Dim objClause As New TestAdioMediaPlayer
    
    TestName = "TestAdioMediaPlayer"
    
    Select Case TestNum
        Case objTestNum.NextNumber: Call TestMediaPlayerLoadFile
        Case objTestNum.NextNumber: Call TestPlayback
        Case objTestNum.NextNumber: Call TestStop
        Case objTestNum.NextNumber: Call TestPause
        Case objTestNum.NextNumber: Call TestSeekForwardAndReserve
        Case objTestNum.NextNumber: Call TestEndOfStream
        
        Case Else
            EndOfTests = True
    End Select
End Sub
Private Sub TestMediaPlayerLoadFile()
With Form_Test.AdioMediaPlayer
    .LoadFile App.Path & "\test.mp3"
    
    AssertEqual 0, Form_Test.MediaPlayerErrorCount, "Loading of a audio file should succeed"
    
    AssertTrue (.State = AdioReady), "State should be `Ready`"
End With
End Sub
Private Sub TestPlayback()
With Form_Test.AdioMediaPlayer
    .SetVolume 0
    .StartPlay
    
    AssertTrue (.State = AdioPlaying), "State should be `Playing`"
    
    AssertEqual 168, .GetProperties.DurationInSeconds, "DurationInSeconds should be 168"
End With
End Sub
Private Sub TestStop()
With Form_Test.AdioMediaPlayer
    .StopPlay
    
    AssertTrue (.State = AdioStopped), "State should be `Stopped`"
End With
End Sub
Private Sub TestPause()
With Form_Test.AdioMediaPlayer
    .PausePlay
    
    AssertTrue (.State = AdioPaused), "State should be `Paused`"
End With
End Sub
Private Sub TestSeekForwardAndReserve()
With Form_Test.AdioMediaPlayer
    .StartPlay
    .SeekBySeconds AdioForward, 20
    
    AssertTrue (.State = AdioPlaying), "State should be `Playing`"
    
    AssertEqual 20, .GetProperties.ElapsedInSeconds, "ElapsedInSeconds should be 20"
    AssertEqual 148, .GetProperties.RemainingInSeconds, "RemainingInSeconds should be 148"
    
    .SeekBySeconds AdioRewind, 10
    
    AssertEqual 10, .GetProperties.ElapsedInSeconds, "ElapsedInSeconds should be 10"
    AssertEqual 158, .GetProperties.RemainingInSeconds, "RemainingInSeconds should be 158"
End With
End Sub
Private Sub TestEndOfStream()
With Form_Test.AdioMediaPlayer
    .SeekBySeconds AdioForward, 157
    
    Do While .State = AdioPlaying
        DoEvents
    Loop
    
    AssertTrue (.State = AdioEnded), "State should be `Ended`"
    AssertEqual 0, .GetProperties.RemainingInSeconds, "RemainingInSeconds should be 0 when stream ended"
End With
End Sub
