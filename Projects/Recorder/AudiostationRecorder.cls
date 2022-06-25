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
Public Function StartRecorder()
Form_Main.btnRecord_Click
End Function
Public Function StopRecorder()
Form_Main.btnRecord_Click
End Function
Public Function SaveRecording()
Form_Main.btnSave_Click
End Function
Public Function Settings()
Form_Settings.Show vbModal
End Function
