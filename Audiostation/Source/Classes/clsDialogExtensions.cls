VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialogExtensions"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////
'// FileName        : DialogExtensions.cls
'// FileType        : Microsoft Visual Basic 6 - Class Module
'// Author          : Alex van den Berg
'// Created         : 26-06-2023
'// Last Modified   : 05-11-2023
'// Copyright       : Sibra-Soft
'// Description     : Sibra-Soft dialog extension class
'////////////////////////////////////////////////////////////////

Option Explicit

Public Enum DlgType
    [dlgNew]
    [dlgEdit]
End Enum

Public DialogType As DlgType
Public Result As VbMsgBoxResult
Public Parameters As New Dictionary
Public Function ShowDialog(Dialog As Form, Optional Owner As Form, Optional Caption As String = vbNullString) As VbMsgBoxResult
If Caption = vbNullString Then: Caption = Space(1)

Dialog.Caption = Caption
Dialog.Show vbModal, Owner

ShowDialog = Result
End Function
