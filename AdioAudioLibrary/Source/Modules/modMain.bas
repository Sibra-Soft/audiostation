Attribute VB_Name = "modMain"
Option Explicit

Public SoundFont As Long

Public TagLib As New clsAdioTagLibrary
Public StringHelpers As New clsStringExtensions
Public Helpers As New clsSibraSoft
Public Function CheckFileSupport(file As String) As Boolean
Dim Fso As New FileSystemObject

' Check if the file is supported by Adio
Select Case LCase(Fso.GetExtensionName(file))
    Case "mp1", "mp2", "mp3", "wav", "ogg", "aiff", "aac", "wma", "flac": CheckFileSupport = True: Exit Function
    Case "mid", "midi", "kar", "rmi", "sid", "mus": CheckFileSupport = True: Exit Function
    
    Case Else: CheckFileSupport = False
End Select
End Function
