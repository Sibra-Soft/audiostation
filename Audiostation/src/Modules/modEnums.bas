Attribute VB_Name = "ModEnums"
Public Enum enumCDMode
    [RandomMode]
    [LoopMode]
End Enum

Public Enum enumPlayStates
    Paused
    Stopped
    Playing
    MediaEnded
End Enum

Public Enum enumPlayMode
    [PlaySingleTrack]
    [AutoNextTrack]
    [Shuffle]
End Enum

Public Enum enumPlaylistMode
    [RepeatPlaylist]
    [RepeatSingleTrack]
End Enum

Public Enum enumMediaMode
    [MidiMediaMode]
    [MP3MediaMode]
    [CDMediaMode]
    [SidMediaMode]
    [MusMediaMode]
End Enum
