Attribute VB_Name = "modMain"
Public Settings As New RegistrySettings
Sub Main()
If App.PrevInstance Then
    Select Case Trim(Command)
        Case "-stop": Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecorderCommand", "stop")
        Case "-save": Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "RecorderCommand", "save")
    End Select
    
    End
Else
    Select Case Trim(Command)
        Case "-settings": Form_Settings.Show
        Case "-record":
            Form_Main.Hide
            Form_Main.btnRecord_Click
            
        Case Else
            End
            
    End Select
End If
End Sub

