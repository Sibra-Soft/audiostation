Attribute VB_Name = "ModBassNetRadio"
Option Explicit

Public Url As Variant
Public TmpNameHold As String
Public TmpNameHold2 As String

Public proxy(100) As Byte ' proxy server

' SAVE LOCAL COPY
Public WriteFile As New FileIO
Public FileIsOpen As Boolean, GotHeader As Boolean
Public DownloadStarted As Boolean, DoDownload As Boolean
Public DlOutput As String, SongNameUpdate As Boolean

' THREADING
Public cthread As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As String, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' MESSAGE BOX
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub DoMeta()
    Dim META As Long
    Dim p As String, tmpMeta As String
    
    META = BASS_ChannelGetTags(chan, BASS_TAG_META)
    
    If META = 0 Then Exit Sub
    tmpMeta = VBStrFromAnsiPtr(META)
    
    If ((Mid(tmpMeta, 1, 13) = "StreamTitle='")) Then
        p = Mid(tmpMeta, 14)
        TmpNameHold = Mid(p, 1, InStr(p, ";") - 2)
        
        AudioStaStreamer.MetaTitle = TmpNameHold
        Form_Main.Label_StreamTitle.Caption = AudioStaStreamer.MetaTitle
        
        If TmpNameHold = TmpNameHold2 Then
            ' do noting
        Else
            TmpNameHold2 = TmpNameHold
            GotHeader = False
            DownloadStarted = False
        End If
        
        DlOutput = App.path & "\" & RemoveSpecialChar(Mid(p, 1, InStr(p, ";") - 2)) & ".mp3"
    End If
End Sub

Sub MetaSync(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    Call DoMeta
End Sub

Sub EndSync(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
    With Form_Main
        .label_StreamStatus.Caption = "Not playing"
    End With
End Sub
' The following functions where added by Peter Hebels
Public Sub SUBDOWNLOADPROC(ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
    If (buffer And length = 0) Then
        'frmNetRadio.lblBPS.Caption = VBStrFromAnsiPtr(buffer) ' display connection status
        Exit Sub
    End If

    If (Not DoDownload) Then
        DownloadStarted = False
        Call WriteFile.CloseFile
        Exit Sub
    End If

    If (Trim(DlOutput) = "") Then Exit Sub

    If (Not DownloadStarted) Then
        DownloadStarted = True
        Call WriteFile.CloseFile
        If (WriteFile.OpenFile(DlOutput)) Then
            SongNameUpdate = False
        Else
            
            SongNameUpdate = True
            
            GotHeader = False
        End If
    End If

    If (Not SongNameUpdate) Then
        If (length) Then
            Call WriteFile.WriteBytes(buffer, length)
        Else
            Call WriteFile.CloseFile
            GotHeader = False
        End If
    Else
        DownloadStarted = False
        Call WriteFile.CloseFile
        GotHeader = False
    End If
End Sub

Public Function RemoveSpecialChar(strFilename As String)
    Dim i As Byte
    Dim SpecialChar As Boolean
    Dim SelChar As String, OutFileName As String

    For i = 1 To Len(strFilename)
        SelChar = Mid(strFilename, i, 1)
        SpecialChar = InStr(":/\?*|<>" & Chr$(34), SelChar) > 0

        If (Not SpecialChar) Then
            OutFileName = OutFileName & SelChar
            SpecialChar = False
        Else
            OutFileName = OutFileName
            SpecialChar = False
        End If
    Next i

    RemoveSpecialChar = OutFileName
End Function
