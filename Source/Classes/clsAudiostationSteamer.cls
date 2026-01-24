VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudiostationSteamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' /////////////////////////////////////////////////////////////////////////////////
' Class:            AudiostationStreamer
' Description:      Adds stream functionality to the Audiostation program
'
' Date Changed:     05-10-2021
' Date Created:     04-10-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public MetaTitle As String
Public Status As String
Public Url As String
Public Name As String
Public Error As Boolean
Public Function OpenStream(Url As String, Optional Name As String = vbNullString) As Boolean
Call BASS_SetConfigPtr(BASS_CONFIG_NET_PROXY, vbNullString)
Call BASS_StreamFree(chan)
Call BASS_ChannelFree(chan)

Me.Error = False
Status = "Connecting..."
MetaTitle = ""

chan = BASS_StreamCreateURL(Url, 0, BASS_STREAM_BLOCK Or BASS_STREAM_STATUS Or BASS_STREAM_AUTOFREE, AddressOf SUBDOWNLOADPROC, 0)

If chan = 0 Then
    Status = "Not playing"
    MsgBox "Can't play the stream"
    
    OpenStream = False
    Me.Error = True
End If

Me.Url = Url
Me.Name = Name

OpenStream = True
End Function
