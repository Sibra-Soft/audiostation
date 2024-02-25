Attribute VB_Name = "modSidPlayer"
Option Explicit
Public Function NewPlayer(FileName As String) As Long
Form_Main.ShellPipe.Run "sidplayer.exe " & Chr(34) & "T:\8-BITS (Commodore 64)\V-GA.sid" & Chr(34), App.path
End Function
Public Sub StopPlayer()

End Sub
Public Sub NextSidSong()

End Sub
Public Sub PreviousSidSong()

End Sub
