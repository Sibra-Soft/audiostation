Attribute VB_Name = "ModMain"
'//////////////////////////////////////////////
' Main application module
' Last changed: 05-10-2021
'//////////////////////////////////////////////

Public Enum enumFormTypes
    [MidiPlayer]
    [Mp3Player]
End Enum

Public LanguageFile As String

' CD Player Class
Public Const MAXDRIVES = 10
Public curdrive As Long
Public stream(MAXDRIVES) As Long
Public seeking As Long

Public IsDebuggig As Boolean
Public Mp3Info As New Mp3Info
Public WebRequest As New WebClient
Public Settings As New RegistrySettings
Public Extensions As New SibraSoft

Public AudioStaStreamer As New AudiostationSteamer
Public Sub Main()
Call ApplicationConstructor
End Sub
Public Sub ApplicationConstructor()
Dim MediaFile As String
Dim MediaIndex As String
Dim MediaDuration As String

If Command = "debugging" Then: IsDebuggig = True

'Check if the temp folder exists
If Dir(App.path & "\temp\", vbDirectory) = vbNullString Then: MkDir (App.path & "\temp\")

'Set the current application langauge
LanguageFile = LCase(Settings.ReadSetting("Sibra-Soft", "Audiostation", "Langauge", "english"))

'Get the loaded file
MediaFile = Command$
MediaFile = Replace(MediaFile, Chr(34), vbNullString)

If Not Extensions.FileExists(MediaFile) Then GoTo Einde
If MediaFile = "" Then: GoTo Einde

If App.PrevInstance Then Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CheckFile", MediaFile): End

Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CheckFile", MediaFile): GoTo Einde
Exit Sub

Einde:
    Form_Main.Show
End Sub
Public Sub ApplicationDestructor()
Call BASS_ChannelFree(chan)
Call BASS_Free
End Sub
