Attribute VB_Name = "modMain"
'///////////////////////////////////////////////////////////////
'// FileName        : modMain.bas
'// FileType        : Microsoft Visual Basic 6 - Module
'// Author          : Alex van den Berg
'// Created         : 04-10-2021
'// Last Modified   : 29-01-2026
'// Copyright       : Sibra-Soft
'// Description     : Main application module
'////////////////////////////////////////////////////////////////

Option Explicit

Public ConfigFile As String

Public IsDebuggig As Boolean

Public StrExt As New clsStringExtensions
Public Extensions As New clsSibraSoft
Public Dialogs As New clsDialogExtensions
Public AppLog As New clsLogger

Public StreamUrl As String
Public StreamName As String

Public CurrentMediaPlayerTrackNr As Long
Public CurrentMidiPlayerTrackNr As Long
Public Sub Main()
Call ApplicationConstructor
End Sub
Public Sub OpenFile(MediaFile As String)
Dim LastTrackAdded As Long
Dim TrackNr As Long

Begin:
If Not Extensions.FileExists(MediaFile) Then Exit Sub

Call AppLog.LogInfo("Load file: " & MediaFile)

With Form_Main
    Select Case LCase(Right(MediaFile, 3))
        Case "mp3", "wav", "mp2", "cda", "wma", "m4a", "ogg" ' Media files
            TrackNr = .AdioMediaPlaylist.AddFile(MediaFile).nR
            Call .AdioMediaPlaylist.GetTrack(PLS_GOTO, TrackNr)
            
        Case "mid", "kar", "mus", "sid" ' Midi files
            TrackNr = .AdioMidiPlaylist.AddFile(MediaFile).nR
            Call .AdioMidiPlaylist.GetTrack(PLS_GOTO, TrackNr)
            
        Case "apl", "wpl", "m3u", "pls" 'Playlist files
            Screen.MousePointer = vbHourglass
            
            Form_Playlist.FormType = Mp3Player
            Select Case LCase(Right(MediaFile, 3))
                Case "apl": Call .AdioMediaPlaylist.LoadPlaylist(MediaFile, PLAYLIST_APL)
                Case "m3u": Call .AdioMediaPlaylist.LoadPlaylist(MediaFile, PLAYLIST_M3U)
                Case "pls": Call .AdioMediaPlaylist.LoadPlaylist(MediaFile, PLAYLIST_PLS)
                Case "wpl": Call .AdioMediaPlaylist.LoadPlaylist(MediaFile, PLAYLIST_WPL)
            End Select
            
            Form_Playlist.Show , Form_Main
    Case Else
        'Check if it's a file that needs to be converted
        Select Case LCase(Right(MediaFile, 3))
            Case "flac"
                TrackNr = .AdioMediaPlaylist.AddFile(MediaFile).nR
                Call .AdioMediaPlaylist.GetTrack(PLS_GOTO, TrackNr)
        End Select
    
        'Check if it's a file that needs to be converted
        Select Case LCase(Right(MediaFile, 2))
            Case "ra": 'Call ModConvert.Convert(MediaFile, [Real Audio], MP3): GoTo Begin
            Case "rm": 'Call ModConvert.Convert(MediaFile, [Real Media], MP3): GoTo Begin
    
            Case Else: MsgBox GetTranslation(1057), vbExclamation
        End Select
    End Select
End With
End Sub
Public Sub ApplicationConstructor()
Dim MediaFile As String

Call argProcessCMDLine

ChDrive App.path
ChDir App.path

' Create folders that are used by the application
Call Extensions.CreateFolderIfNotExists(Environ$("AppData") & "\Audiostation")
Call Extensions.CreateFolderIfNotExists(Environ$("AppData") & "\Audiostation\temp\")
Call Extensions.CreateFolderIfNotExists(Environ$("AppData") & "\Audiostation\logs\")

' Copy settings file
If Not Extensions.FileExists((Environ$("AppData") & "\Audiostation\settings.ini")) Then
    Call FileCopy(App.path & "\settings.ini", (Environ$("AppData") & "\Audiostation\settings.ini"))
End If

' Set the folder for the application logfiles
Call AppLog.Init(Environ$("AppData") & "\Audiostation\logs\")

' Get the settings file
ConfigFile = Environ$("AppData") & "\Audiostation\settings.ini"
Call AppLog.LogInfo("Application config file set: " & ConfigFile)

' Set the current application langauge
Call SetLanguage(Extensions.INIRead("main", "Langauge", ConfigFile, "english"))

' Get the loaded file
If UBound(modArgs.argv) > 0 Then: MediaFile = modArgs.argv(1)
If Not Extensions.FileExists(MediaFile) Then GoTo Einde
If MediaFile = "" Then: GoTo Einde

' Make sure to load the file in the last instance of the application
If App.PrevInstance Then
    Form_System.DataInter.Connect 15448
    Form_System.DataInter.SendData "OpenFile~" & MediaFile
    End
Else
    Call OpenFile(MediaFile)
    GoTo Einde
End If

Exit Sub
Einde:
    Form_Main.Show
End Sub
