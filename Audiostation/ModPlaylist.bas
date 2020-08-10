Attribute VB_Name = "ModPlaylist"
Public Function OpenAplPlaylist(strPlaylistFile As String)
Dim Files As String

Files = Extensions.FileGetContents(strPlaylistFile)

Call Form_Playlist.AddToPlaylist(Files)
End Function
Public Function OpenWplPlaylist(FileName As String)
Dim Lines
Dim FileContent As String
Dim I As Integer
Dim Media As String
Dim Files As String

FileContent = Extensions.FileGetContents(FileName)
Lines = Split(FileContent, vbNewLine)

For I = 0 To UBound(Lines)
    If InStr(1, Lines(I), "<media") Then
        Media = Extensions.StringBetween("<media", "/>", Trim(Lines(I)))
        Media = Replace(Media, Chr(34), vbNullString)
        Media = Replace(Media, "media src=", vbNullString)
        
        Files = Files & Media & vbNewLine
    End If
Next

Call Form_Playlist.AddToPlaylist(Files)
End Function
Public Function OpenM3uPlaylist(strPlaylistFile As String, Optional TargetListbox As ListBox)
Dim TextLine As String, FN As Integer
Dim Files As String

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
                    Files = Files & TextLine & vbNewLine
                End If
            End If
        End If
    Loop
Close #FN

'Add all the files from the array to the playlist
Call Form_Playlist.AddToPlaylist(Files)
End Function
Public Function OpenPlsPlaylist(strPlaylistFile As String, Optional TargetListbox As ListBox)
Dim I As Integer
Dim strNumberofEntries As Integer
Dim FileToAdd As String

strNumberofEntries = Extensions.INIRead("playlist", "NumberOfEntries", strPlaylistFile)

For I = 1 To strNumberofEntries
    FileToAdd = Extensions.INIRead("playlist", "File" & I, strPlaylistFile)

    Call Form_Playlist.AddToPlaylist(FileToAdd)
Next
End Function
