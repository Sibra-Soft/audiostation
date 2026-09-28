Attribute VB_Name = "modMidiUtils"
Option Explicit

Public gisEnd As Boolean
Public gisCurrentDoEvents As Boolean
Public gisCurrentQueue As Boolean
Public gisCurrentFF As Boolean
Public gmThreadPriorityApp As Integer

Public Type MidiFFTracking
    nDetectedMessageNumber As Long
End Type

Public Const MB_INTEGERUBOUND = 32767
Public Const MB_LONGUBOUND = &H7FFFFFFF
Public Const MB_LOWNIBBLE = &HF
Public Const MB_HIGHNIBBLE = &HF0
Public Const MB_LOWBYTE = &HFF
Public Const MB_HIGHBYTE = &HFF00
Public Const MB_DOEVENTSPOLLING = 10
Public Const MB_DEVICEID = &H10
