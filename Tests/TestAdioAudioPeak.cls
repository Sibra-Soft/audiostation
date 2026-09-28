VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdioAudioPeak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdioLibrary - AdioAudioPeak Component
'
Option Explicit

Implements ITestCase
Private Function ITestCase_TestCaseInstance() As ITestCase
    Dim objResult As New TestAdioAudioPeak
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
    Dim objClause As New TestAdioAudioPeak
    
    TestName = "TestAdioAudioPeak"
    
    Select Case TestNum
        Case objTestNum.NextNumber: Call TestComponentInitialization
        
        Case Else
            EndOfTests = True
    End Select
End Sub
Private Sub TestComponentInitialization()
With Form_Test
    .AdioAudioPeak.Run
End With
End Sub


