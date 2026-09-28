Attribute VB_Name = "modFileProp"
Private Function FolderName(FileName As String) As String
Dim posn As Integer

posn = InStrRev(FileName, "\")

If posn > 0 Then
    FolderName = Left$(FileName, posn)
Else
    FolderName = ""
End If
End Function
Public Function GetFileDurationInSeconds(file As String) As Long
Dim Hours, Minutes, seconds As Integer
Dim TotalSeconds As Long
Dim Fso As New FileSystemObject

Dim objShell As Object
Dim objFolder, oFile As Object
Dim time() As String

Set objFolder = Nothing
Set objShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(FolderName(file))
Set oFile = objFolder.ParseName(Fso.GetFileName(file))

If Not oFile Is Nothing Then
    Call modOS.GetWindowsVersion
    If modOS.VerMarjor >= 6 Then
        time = Split(objFolder.GetDetailsOf(oFile, 27), ":")
    Else
        time = Split(objFolder.GetDetailsOf(oFile, 21), ":") ' Windows XP
    End If
    
    If UBound(time) = 2 Then
        Hours = time(0)
        Minutes = time(1)
        seconds = time(2)
        
        TotalSeconds = TotalSeconds + Hours * 60 * 60
        TotalSeconds = TotalSeconds + Minutes * 60
        TotalSeconds = TotalSeconds + seconds
    End If
End If

GetFileDurationInSeconds = TotalSeconds
End Function
