VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudiostationRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Class:            AudiostationRecorder
' Description:      Adds recorder functionality to the Audiostation program
'
' Date Changed:     25-06-2022
' Date Created:     25-06-2022
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public Recording As Boolean
Public Sub Settings()
Call Shell(App.path & "\recorder.exe -settings", vbNormalFocus)
End Sub
Public Sub StartRecorder()
Call Shell(App.path & "\recorder.exe -record", vbNormalFocus)
Recording = True
End Sub
Public Sub StopRecorder()
Call Shell(App.path & "\recorder.exe -stop", vbNormalFocus)
Recording = False
End Sub
Public Sub SaveRecording()
StopRecorder
Call Shell(App.path & "\recorder.exe -save", vbNormalFocus)
End Sub
