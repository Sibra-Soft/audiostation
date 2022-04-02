Attribute VB_Name = "ModLanguage"
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Module:           ModLanguage
' Description:      Various functions for multi-language applications
'
' Date Changed:     25-10-2021
' Date Created:     04-10-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public Function SetLanguage(Frm As Form)
Dim ControlCount As Integer
Dim ControlTranslateCount As Integer
Dim ControlMenuTranslateCount As Integer
Dim CurrentControl As Control

On Error Resume Next
If Not Frm.Tag = vbNullString Then: Frm.Caption = GetLanguage(Frm.Tag)

For Each CurrentControl In Frm.Controls
    If Not CurrentControl.Tag = vbNullString Then
        CurrentControl.Caption = GetLanguage(CurrentControl.Tag)
        ControlTranslateCount = ControlTranslateCount + 1
    End If
    
    If Not CurrentControl.HelpContextID = 0 Then
        CurrentControl.Caption = GetLanguage(CurrentControl.HelpContextID)
        ControlMenuTranslateCount = ControlMenuTranslateCount + 1
    End If
    
    ControlCount = ControlCount + 1
Next

Debug.Print "Found " & ControlCount & " control(s), translated " & ControlTranslateCount & " control(s) and " & ControlMenuTranslateCount & " menu item(s)"
End Function
Public Function GetLanguage(TextID As Integer) As String
GetLanguage = Extensions.INIRead("language", str(TextID), App.path & "\languages\" & LanguageFile & ".lng")
GetLanguage = Replace(GetLanguage, "\n", vbNewLine)
End Function
