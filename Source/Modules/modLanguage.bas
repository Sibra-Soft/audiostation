Attribute VB_Name = "modLanguage"
'///////////////////////////////////////////////////////////////
'// FileName        : modLanguage.bas
'// FileType        : Microsoft Visual Basic 6 - Module
'// Author          : Alex van den Berg
'// Created         : 26-06-2023
'// Last Modified   : 05-11-2023
'// Copyright       : Sibra-Soft
'// Description     : Multi-language application module
'////////////////////////////////////////////////////////////////

Option Explicit

Private Language As String

Public File As String
Public Sub TranslateFormAndControls(Frm As Form, Optional UseMnemonic As Boolean)
Dim MenuCount, ControlCount As Integer
Dim CurrentControl As Control

MenuCount = 0
ControlCount = 0

If Not StrExt.IsNullOrWhiteSpace(Frm.Tag) Then
    If UseMnemonic Then
        Frm.Caption = GetTranslation(Frm.Tag)
    Else
        Frm.Caption = Replace(GetTranslation(Frm.Tag), "&", vbNullString)
    End If
End If

On Error Resume Next
For Each CurrentControl In Frm.Controls
    If StrExt.StartsWith("menu", CurrentControl.Name, False) Then
        If Not CurrentControl.HelpContextID = 0 Then
            If UseMnemonic Then
                CurrentControl.Caption = GetTranslation(CurrentControl.HelpContextID)
            Else
                CurrentControl.Caption = Replace(GetTranslation(CurrentControl.HelpContextID), "&", vbNullString)
            End If
            
            MenuCount = MenuCount + 1
        End If
    End If
    
    If InStr(1, CurrentControl.Caption, "T(") > 0 Then
        Dim TranslationId As Integer
        
        TranslationId = Replace(Replace(CurrentControl.Caption, ")", vbNullString), "T(", vbNullString)
        CurrentControl.Caption = GetTranslation(TranslationId)
        ControlCount = ControlCount + 1
    End If
Next

Call AppLog.LogInfo("TranslateFormAndControls: Done - Menu's: " & MenuCount & " - Controls: " & ControlCount)
End Sub
Public Function SetLanguage(Language As String)
File = App.Path & "\languages\" & Language & ".lng"
Language = Language

If Not Extensions.FileExists(File) Then MsgBox "The specified language file could not be found: " & vbNewLine & modLanguage.File, vbOKOnly + vbExclamation, "Error": End

Call AppLog.LogInfo("SetLanguage: " & File & " - " & Language)
End Function
Public Function GetTranslation(TranslationId As Integer, Optional UseMnemonic As Boolean = True) As String
GetTranslation = Extensions.INIRead("language", Str(TranslationId), File)

If Not UseMnemonic Then: GetTranslation = Replace(GetTranslation, "&", vbNullString)

GetTranslation = Replace(GetTranslation, "\n", vbNewLine)
End Function
