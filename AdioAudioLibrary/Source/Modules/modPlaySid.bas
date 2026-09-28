Attribute VB_Name = "modPlaySid"
Option Explicit

Private hMain As Long
Public Function LoadFile(strFile As String) As Boolean
Dim AppResult As Boolean

AppResult = StartExternalApp(App.path & "\sidplay2w.exe", Chr(34) & strFile & Chr(34), True)

If Not AppResult Then
    MsgBox "Starten van de applicatie is mislukt."
    Exit Function
End If

hMain = WaitForMainWindowByPID(g_pi.dwProcessId, 10000, False)

If hMain = 0 Then
    MsgBox "Hoofdvenster niet gevonden."
    Exit Function
End If

LoadFile = True
End Function
Public Sub PausePlay()
Call ClickButtonInWindow(hMain, "Pause")
End Sub
Public Sub StartPlay()
Call ClickButtonInWindow(hMain, "Play")
End Sub
Public Sub StopPlay()
Call ClickButtonInWindow(hMain, "Stop")

Call CleanupExternalApp
Call CloseExternalApp(hMain)
End Sub
Public Sub NextSidTrack()
Dim hScroll As Long

hScroll = FindFirstScrollBar(hMain)

If hScroll = 0 Then Exit Sub

Call ScrollBarStepForward(hScroll)
End Sub
Public Sub PreviousSidTrack()
Dim hScroll As Long

hScroll = FindFirstScrollBar(hMain)

If hScroll = 0 Then Exit Sub

Call ScrollBarStepBack(hScroll)
End Sub
