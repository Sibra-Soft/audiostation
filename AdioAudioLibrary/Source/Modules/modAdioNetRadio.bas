Attribute VB_Name = "modAdioNetRadio"
'///////////////////////////////////////////////////////////////
'// FileName        : modAdioNetRadio.bas
'// FileType        : Microsoft Visual Basic 6 - Module
'// Author          : Alex van den Berg
'// Created         : 10-10-2023
'// Last Modified   : 15-10-2023
'// Copyright       : Sibra-Soft
'// Description     : Simplified functions for net radio streaming
'////////////////////////////////////////////////////////////////

Option Explicit

'// Enums
Public Enum enumAdioNetRadioState
    [Buffering]
    [Playing]
    [Stopped]
End Enum

'// Public vars
Public StreamMeta As String
Public StreamState As enumAdioNetRadioState
Public StreamBufferProgress As Integer

'// Private vars
Private chan As Long
Private Req As Long
Private Const BASS_SYNC_HLS_SEGMENT = &H10300
Private Const BASS_TAG_HLS_EXTINF = &H14000
'*
'* Read meta data of the current specified stream
'*
Public Sub DoMeta()
Dim meta As Long
Dim metaTxt As String

meta = BASS_ChannelGetTags(chan, BASS_TAG_META)

If meta Then ' got Shoutcast metadata
    metaTxt = VBStrFromAnsiPtr(meta)
    Dim p As Long
    p = InStr(metaTxt, "StreamTitle='") ' locate the title
    If p Then
        Dim p2 As Long
        p2 = InStr(p, metaTxt, "';") ' locate the end of it
        If p2 Then
            StreamMeta = mId(metaTxt, p + 13, p2 - (p + 13))
        End If
    End If
Else
    meta = BASS_ChannelGetTags(chan, BASS_TAG_OGG)
    Dim Artist As String, Title As String
    If meta Then ' got Icecast/OGG tags
        Do
            metaTxt = VBStrFromAnsiPtr(meta)
            If metaTxt = vbNullString Then Exit Do
            If Left(metaTxt, 7) = "artist=" Then ' found the artist
                Artist = mId(metaTxt, 8)
            ElseIf Left(metaTxt, 6) = "title=" Then ' found the title
                Artist = mId(metaTxt, 7)
            End If
            meta = meta + Len(metaTxt) + 1
        Loop
        If Title <> "" Then
            If Artist <> "" Then
                StreamMeta = Artist & " - " & Title
            Else
                StreamMeta = Title
            End If
        End If
    Else
        meta = BASS_ChannelGetTags(chan, BASS_TAG_HLS_EXTINF)
        If meta Then ' got HLS segment info
            metaTxt = VBStrFromAnsiPtr(meta)
            p = InStr(meta, ",")
            If p Then StreamMeta = mId(metaTxt, p + 1)
        End If
    End If
End If
End Sub
'*
'* Open a URL as stream
'* @param String StreamUrl: The URL of the stream you want to open
'* @param String ProxyServer: Optional proxy server address
'* @return Boolean: Tells if the stream could be opened
'*
Public Function OpenStreamByUrl(StreamUrl As String, Optional ProxyServer As String = vbNullString) As Boolean
Dim R As Long

' Enable playlist processing
Call BASS_SetConfig(BASS_CONFIG_NET_PLAYLIST, 1)

' Check if the proxy must be set
If StringHelpers.IsNullOrWhiteSpace(ProxyServer) Then
    Call BASS_SetConfigPtr(BASS_CONFIG_NET_PROXY, vbNullString)
Else
    Call BASS_SetConfigPtr(BASS_CONFIG_NET_PROXY, StrPtr(ProxyServer))
End If

Req = Req + 1
R = Req

If chan Then BASS_StreamFree chan ' close old stream

chan = BASS_StreamCreateURL(StreamUrl, 0, BASS_STREAM_BLOCK Or BASS_STREAM_STATUS Or BASS_STREAM_AUTOFREE Or BASS_SAMPLE_FLOAT, AddressOf STATUSPROC, R) ' open URL

If R <> Req Then ' there is a newer request, discard this stream
    If chan Then Call BASS_StreamFree(chan)
    Exit Function
End If

If chan = 0 Then ' failed to open
    OpenStreamByUrl = False
Else
    Dim proc As Long
    
    Call BASS_ChannelSetAttributeEx(chan, BASS_ATTRIB_DOWNLOADPROC, proc, LenB(proc))

    Call BASS_ChannelSetSync(chan, BASS_SYNC_META, 0, AddressOf METASYNC, 0) ' Shoutcast
    Call BASS_ChannelSetSync(chan, BASS_SYNC_OGG_CHANGE, 0, AddressOf METASYNC, 0) ' Icecast/OGG
    Call BASS_ChannelSetSync(chan, BASS_SYNC_HLS_SEGMENT, 0, AddressOf METASYNC, 0) ' HLS
    Call BASS_ChannelSetSync(chan, BASS_SYNC_STALL, 0, AddressOf STALLSYNC, 0)
    Call BASS_ChannelSetSync(chan, BASS_SYNC_FREE, 0, AddressOf FREESYNC, 0)
    
    Call BASS_ChannelPlay(chan, BASSFALSE)
    
    OpenStreamByUrl = True
    
    StreamState = Playing
End If
End Function

Public Sub TimerProc()
Dim active As Long

active = BASS_ChannelIsActive(chan)

If active = BASS_ACTIVE_STALLED Then
    StreamState = Buffering
    StreamBufferProgress = (100 - Int(BASS_StreamGetFilePosition(chan, BASS_FILEPOS_BUFFERING)))
Else
    If active Then
        StreamState = Playing
        Dim icy As Long
        
        icy = BASS_ChannelGetTags(chan, BASS_TAG_ICY)
        
        If icy = 0 Then
            icy = BASS_ChannelGetTags(chan, BASS_TAG_HTTP) ' no ICY tags, try HTTP
            If icy Then
                Dim icyTxt As String
                Do
                    icyTxt = VBStrFromAnsiPtr(icy)
                    
                    If icyTxt = "" Then Exit Do
                    If Left(icyTxt, 9) = "icy-name:" Then
                        StreamMeta = mId(icyTxt, 10)
                    ElseIf Left(icyTxt, 8) = "icy-url:" Then
                        StreamMeta = mId(icyTxt, 9)
                    End If
                    
                    icy = icy + Len(icyTxt) + 1
                Loop
            End If
        End If
        
        Call DoMeta
    End If
End If
End Sub
Public Sub METASYNC(ByVal handle As Long, ByVal Channel As Long, ByVal data As Long, ByVal user As Long)
Call DoMeta
End Sub

Public Sub STALLSYNC(ByVal handle As Long, ByVal Channel As Long, ByVal data As Long, ByVal user As Long)
If data = 0 Then ' stalled
    'frmNetradio.tmrStall.Enabled = True ' start buffer monitoring
End If
End Sub

Public Sub FREESYNC(ByVal handle As Long, ByVal Channel As Long, ByVal data As Long, ByVal user As Long)
chan = 0

StreamState = Stopped
End Sub

Public Sub STATUSPROC(ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
If (buffer <> 0) And (length = 0) And (user = Req) Then
    'frmNetradio.lbl32.Caption = VBStrFromAnsiPtr(buffer) ' display status
End If
End Sub

