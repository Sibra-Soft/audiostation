VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mdlAdioProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'///////////////////////////////////////////////////////////////
'// FileName        : mdlAdioProperties.cls
'// FileType        : Microsoft Visual Basic 6 - Class Module
'// Type            : Model
'// Author          : Alex van den Berg
'// Created         : 30-01-2026
'// Last Modified   : 28-09-2026
'// Copyright       : Sibra-Soft
'// Description     : Adio audio file properties
'////////////////////////////////////////////////////////////////

Public DurationString As String
Public DurationInSeconds As Integer

Public ElapsedString As String
Public ElapsedInSeconds As Integer

Public RemainingString As String
Public RemainingInSeconds As Integer

