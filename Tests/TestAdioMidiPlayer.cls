VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdioMidiPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdioLibrary - MidiPlayer Component
'
Option Explicit

Implements ITestCase
Private Function ITestCase_TestCaseInstance() As ITestCase
    Dim objResult As New TestAdioMidiPlayer
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
    Dim objClause As New TestAdioMidiPlayer
    
    TestName = "TestAdioMidiPlayer"
    
    Select Case TestNum
        Case tstcTestActionSetup: Call Setup
    
        Case objTestNum.NextNumber: Call TestComponentInitialization
        Case objTestNum.NextNumber: Call TestMidiGetDeviceList
        Case objTestNum.NextNumber: Call TestMidiPlayerLoadMidiFile
        Case objTestNum.NextNumber: Call TestMidiPlayerLoadKarFile
        
        Case Else
            EndOfTests = True
    End Select
End Sub
Private Sub Setup()
Form_Test.MidiPlayerErrorCount = 0
End Sub
Private Sub TestComponentInitialization()
Dim deviceId As Integer

deviceId = 0

With Form_Test.AdioMidiPlayer
    .InitComponent deviceId
End With
End Sub
Private Sub TestMidiGetDeviceList()
With Form_Test.AdioMidiPlayer
    AssertTrue (.GetListOfMidiDevices.Count > 0), "Atleast one midi device should be available"
End With
End Sub
Private Sub TestMidiPlayerLoadMidiFile()
With Form_Test.AdioMidiPlayer
    .LoadFile App.Path & "\test.mid"

    AssertEqual 0, Form_Test.MidiPlayerErrorCount, "Loading of a midi file should succeed"
    AssertEqual [MID - Midi File], .LoadedFileType, "The file type should be midi after it is loaded"
    
    AssertTrue (.State = AdioReady), "State should be `Ready`"
End With
End Sub
Private Sub TestMidiPlayerLoadKarFile()
With Form_Test.AdioMidiPlayer
    .LoadFile App.Path & "\test.kar"
    
    AssertEqual 0, Form_Test.MidiPlayerErrorCount, "Loading of a kar file should succeed"
    AssertEqual [KAR - Karaoke File], .LoadedFileType, "The file type should be kar after it is loaded"
    AssertEqual 24, .NumberOfTracks, "The test file should contain 24 midi tracks"
    
    AssertTrue (.State = AdioReady), "State should be `Ready`"
End With
End Sub
