VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mdlAdioCdDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'///////////////////////////////////////////////////////////////
'// FileName        : AdioCdDrive.cls
'// FileType        : Microsoft Visual Basic 6 - Class Module
'// Type            : Model
'// Author          : Alex van den Berg
'// Created         : 18-08-2023
'// Last Modified   : 14-10-2023
'// Copyright       : Sibra-Soft
'// Description     : Cd device model
'////////////////////////////////////////////////////////////////

Public cdLetter As String
Public cdDescription As String
