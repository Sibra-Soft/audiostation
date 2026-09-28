Attribute VB_Name = "modGlobals"
Option Explicit

' Declares
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Consts
Public Const ADIO_PLAYSTATE_STOPPED = "AdioStopped"
Public Const ADIO_PLAYSTATE_PLAYING = "AdioPlaying"
Public Const ADIO_PLAYSTATE_PAUSED = "AdioPaused"
Public Const ADIO_PLAYSTATE_ENDED = "AdioEnded"

Public Const ADIO_SEEK_DIRECTION_FORWARD = "AdioForward"
Public Const ADIO_SEEK_DIRECTION_REWIND = "AdioRewind"

Public Const ADIO_FADE_IN = "AdioIn"
Public Const ADIO_FADE_OUT = "AdioOut"
Public Const ADIO_FADE_CROSS = "AdioCross"

' Types
Public Type ADIO_DEVICE
    Name As String
End Type

Public Type ADIO_PROPERTIES
    DurationInSeconds As Integer
    DurationString As String
    ElapsedInSeconds As Integer
    ElapsedString As String
    RemainingInSeconds As Integer
    RemainingString As String
End Type

Public Type ADIO_TAG
    Artist As String
    Title As String
    Album As String
    Year As Integer
    RuntimeInSeconds As Integer
    RuntimeString As String
End Type

Public Type ADIO_PLAYLIST_ITEM
    TrackNr As Integer
    Media As ADIO_TAG
End Type
