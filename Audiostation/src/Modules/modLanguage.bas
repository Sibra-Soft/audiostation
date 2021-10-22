Attribute VB_Name = "ModLanguage"
Public Function SetLanguage(Frm As Form)
Dim ControlCount As Integer
Dim ControlTranslateCount As Integer
Dim ControlMenuTranslateCount As Integer

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
GetLanguage = Extensions.INIRead("language", Str(TextID), App.path & "\languages\" & LanguageFile & ".lng")
GetLanguage = Replace(GetLanguage, "\n", vbNewLine)
End Function
