Attribute VB_Name = "modPlaylist"
Public Function SavePlsPlaylist(file As String, list As Collection) As Boolean
Dim plsItem As mdlAdioPlaylistItem

For Each plsItem In list
    Call Helpers.INIWrite("playlist", "File" & I, plsItem.plsLocalFile, file)
Next

Call Extensions.INIWrite("playlist", "NumberOfEntries", lstFormList.ListItems.Count, file)
Call Extensions.INIWrite("playlist", "Version", 2, file)

' Check if the playlist has been saved
If Helpers.FileExists(file) Then
    SavePlsPlaylist = True
Else
    SavePlsPlaylist = False
End If
End Function
Public Function SaveAplPlaylist(strFile As String, colList As Collection) As Boolean
Dim plsItem As mdlAdioPlaylistItem
Dim FN As Integer

FN = FreeFile

Open strFile For Output As #FN
    For Each plsItem In colList
        Print #FN, plsItem.plsLocalFile
    Next
Close #FN

' Check if the playlist has been saved
If Helpers.FileExists(strFile) Then
    SaveAplPlaylist = True
Else
    SaveAplPlaylist = False
End If
End Function
Public Function SaveM3uPlaylist(file As String, list As Collection) As Boolean
Dim plsItem As mdlAdioPlaylistItem

Open file For Output As #FN
    Print #FN, "#EXTM3U"
    
    For Each plsItem In list
      Print #FN, "#EXTINF:0, " & Helpers.GetFileNameFromFilePath(plsItem.plsLocalFile, False)
      Print #FN, plsItem.plsLocalFile
      Print #FN, ""
    Next
Close #FN

' Check if the playlist has been saved
If Helpers.FileExists(file) Then
    SaveM3uPlaylist = True
Else
    SaveM3uPlaylist = False
End If
End Function
Public Function SaveWplPlaylist(file As String, list As Collection) As Boolean
Dim PlaylistName As String
Dim plsItem As mdlAdioPlaylistItem

PlaylistName = Helpers.GetFileNameFromFilePath(file, False)

Open file For Output As #1
    Print #1, "<?wpl version="; 1#; "?>"
    Print #1, "<smil>"
    Print #1, "    <head>"
    Print #1, "        <title>" & PlaylistName & "</title>"
    Print #1, "    </head>"
    Print #1, "    <body>"
    Print #1, "        <seq>"
    
    ' Get all the items from the selected playlist
    For Each plstItem In list
      Print #1, "<media src=""" & plsItem.plsLocalFile & """/>"
    Next
    
    Print #1, "       </seq>"
    Print #1, "    </body>"
    Print #1, "</smil>"
Close #1

' Check if the playlist has been saved
If Ext.FileExists(file) Then
    SaveWplPlaylist = True
Else
    SaveWplPlaylist = False
End If
End Function
Public Function LoadAplFile(file As String) As String
Dim StringListOfFiles As String
StringListOfFiles = Helpers.FileGetContents(file)

LoadAplFile = StringListOfFiles
End Function
Public Function LoadWplFile(file As String) As String
Dim Lines
Dim FileContent As String
Dim I As Integer
Dim Media As String
Dim StringListOfFiles As String

FileContent = Extensions.FileGetContents(FileName)
Lines = Split(FileContent, vbNewLine)

For I = 0 To UBound(Lines)
    If InStr(1, Lines(I), "<media") Then
        Media = StrExt.Between("<media", "/>", Trim(Lines(I)))
        Media = Replace(Media, Chr(34), vbNullString)
        Media = Replace(Media, "media src=", vbNullString)
        
        StringListOfFiles = StringListOfFiles & Media & vbNewLine
    End If
Next

LoadWplFile = StringListOfFiles
End Function
Public Function LoadM3uFile(file As String) As String
Dim TextLine As String, FN As Integer
Dim StringListOfFiles As String

FN = FreeFile

'Add the files to the array
Open strPlaylistFile For Input As #FN
    Do While Not EOF(FN)
        Line Input #FN, TextLine
        If TextLine <> LineToRem Then
            If Left(TextLine, 7) = "#EXTM3U" Then
                Debug.Print "Playlist Type: M3U"
            Else
                If Left(TextLine, 8) = "#EXTINF:" Then
                    Debug.Print "Info Data: " & TextLine
                Else
                    StringListOfFiles = StringListOfFiles & TextLine & vbNewLine
                End If
            End If
        End If
    Loop
Close #FN

LoadM3uFile = StringListOfFiles
End Function
Public Function LoadPlsFile(file As String) As String
Dim I As Integer
Dim strNumberofEntries As Integer
Dim StringListOfFiles As String

strNumberofEntries = Extensions.INIRead("playlist", "NumberOfEntries", strPlaylistFile)

For I = 1 To strNumberofEntries
    StringListOfFiles = StringListOfFiles & Extensions.INIRead("playlist", "File" & I, strPlaylistFile) & vbNewLine
Next

LoadPlsFile = StringListOfFiles
End Function
