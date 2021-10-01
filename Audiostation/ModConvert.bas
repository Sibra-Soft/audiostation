Attribute VB_Name = "ModConvert"
Public Enum ConvertFrom
    [Real Media]
    [Real Audio]
    [Voice File Format]
    [Apple Core Format]
    [Sony Wave64]
    [OGG]
    [Westwood Studios Audio]
    [Sony OpenMG Audio]
    [Raw Flac]
    [Creative Voice]
    [Windows Media]
    [Media4A]
End Enum

Public Enum convertto
    [MP3]
    [WAV]
End Enum

Public Function Convert(ConvFilename As String, ConvFrom As ConvertFrom, convTo As convertto)
Dim CurrentFilename As String

CurrentFilename = Extensions.GetFileNameFromFilePath(ConvFilename, False)

Call Extensions.ShellAndWait("ffmpeg.exe", "-i " & Chr(34) & ConvFilename & Chr(34) & " " & App.path & "\temp\" & CurrentFilename & ".mp3")
Call Settings.WriteSetting("Sibra-Soft", "Audiostation", "CheckFile", App.path & "\temp\" & CurrentFilename & ".mp3")
End Function
