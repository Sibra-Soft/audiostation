VERSION 5.00
Begin VB.UserControl AdioTagging 
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   3225
   Begin VB.Image Image_Main 
      Height          =   480
      Left            =   0
      Picture         =   "AdioTagging.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "AdioTagging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'// Enums
Public Enum enumAdioTagVersion
    [v1]
    [v2]
End Enum

Public Enum enumAdioProperty
    [pFileSize]
    [pDurationInSeconds]
    [pDurationString]
    [pBitrate]
    [pChannels]
    [pStereo]
    [pFrequency]
End Enum

Public Enum enumAdioTags
    [tArtist]
    [tTitle]
    [tAlbum]
    [tYear]
    [tGenre]
    [tTrackNumber]
    [tComposer]
    [tCopyright]
    [tAlbumArtist]
End Enum
'*
'* Write MP3 meta tags to a specified MP3 file
'* @param String File: The file you want to change
'* @param String Value: The value you want to write
'* @param enumAdioTags TagType: The tag you want to change
'* @param enumAdioTagVersion AdioTags: The version of the tag you want to write (V1, V2)
'* @return Boolean: Tells if the tags has been changed
'*
Public Function WriteTag(file As String, Value As String, TagType As enumAdioTags, TagVersion As enumAdioTagVersion) As Boolean
WriteTag = modAdio.AdioWriteTag(file, Value, TagType, TagVersion)
End Function
'*
'* Get file properties of a specified file
'* @param String File: The file to get the property of
'* @param enumAdioProperty PropertyType: The property you want to get
'* @return String: The value of the property
'*
Public Function ReadProperty(file As String, PropertyType As enumAdioProperty) As String
ReadProperty = modAdio.AdioReadAudioProperty(file, PropertyType)
End Function
'*
'* Gets MP3 meta tags of a specified file
'* @param String File: The file you want to get tags of
'* @param enumAdioTags TagType: The tag you want to get
'* @param enumAdioTagVersion TagVersion: The version of the tag you want to get (V1, V2)
'* @return String: The value of the tag
'*
Public Function ReadTag(file As String, TagType As enumAdioTags, TagVersion As enumAdioTagVersion) As String
ReadTag = modAdio.AdioReadTag(file, TagType, TagVersion)
End Function
'*
'* Resize the usercontrol
'*
Private Sub UserControl_Resize()
width = Image_Main.width
height = Image_Main.height
End Sub
