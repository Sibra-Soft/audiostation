VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mdlAdioMidiDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'///////////////////////////////////////////////////////////////
'// FileName        : mdlAdioMidiDevice.cls
'// FileType        : Microsoft Visual Basic 6 - Class Module
'// Type            : Model
'// Author          : Alex van den Berg
'// Created         : 18-03-2026
'// Last Modified   : 28-09-2026
'// Copyright       : Sibra-Soft
'// Description     : Adio Midi device model
'////////////////////////////////////////////////////////////////

Public mId As Integer
Public mName As String
