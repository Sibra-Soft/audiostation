Attribute VB_Name = "modBassMix"
' BASSmix 2.4 Visual Basic module
' Copyright (c) 2005-2021 Un4seen Developments Ltd.
'
' See the BASSMIX.CHM file for more detailed documentation

' Additional BASS_SetConfig options
Global Const BASS_CONFIG_MIXER_BUFFER = &H10601
Global Const BASS_CONFIG_MIXER_POSEX = &H10602
Global Const BASS_CONFIG_SPLIT_BUFFER = &H10610

' BASS_Mixer_StreamCreate flags
Global Const BASS_MIXER_END = &H10000      ' end the stream when there are no sources
Global Const BASS_MIXER_NONSTOP = &H20000  ' don't stall when there are no sources
Global Const BASS_MIXER_RESUME = &H1000    ' resume stalled immediately upon new/unpaused source
Global Const BASS_MIXER_POSEX = &H2000     ' enable BASS_Mixer_ChannelGetPositionEx support

' BASS_Mixer_StreamAddChannel/Ex flags
Global Const BASS_MIXER_CHAN_ABSOLUTE = &H1000  ' start is an absolute position
Global Const BASS_MIXER_CHAN_BUFFER = &H2000    ' buffer data for BASS_Mixer_ChannelGetData/Level
Global Const BASS_MIXER_CHAN_LIMIT = &H4000     ' limit mixer processing to the amount available from this source
Global Const BASS_MIXER_CHAN_MATRIX = &H10000   ' matrix mixing
Global Const BASS_MIXER_CHAN_PAUSE = &H20000    ' don't process the source
Global Const BASS_MIXER_CHAN_DOWNMIX = &H400000 ' downmix to stereo/mono
Global Const BASS_MIXER_CHAN_NORAMPIN = &H800000 ' don't ramp-in the start
Global Const BASS_MIXER_BUFFER = BASS_MIXER_CHAN_BUFFER
Global Const BASS_MIXER_LIMIT = BASS_MIXER_CHAN_LIMIT
Global Const BASS_MIXER_MATRIX = BASS_MIXER_CHAN_MATRIX
Global Const BASS_MIXER_PAUSE = BASS_MIXER_CHAN_PAUSE
Global Const BASS_MIXER_DOWNMIX = BASS_MIXER_CHAN_DOWNMIX
Global Const BASS_MIXER_NORAMPIN = BASS_MIXER_CHAN_NORAMPIN

' Mixer attributes
Global Const BASS_ATTRIB_MIXER_LATENCY = &H15000
Global Const BASS_ATTRIB_MIXER_THREADS = &H15001

' Additional BASS_Mixer_ChannelIsActive return values
Global Const BASS_ACTIVE_WAITING = 5

' BASS_Split_StreamCreate flags
Global Const BASS_SPLIT_SLAVE = &H1000   ' only read buffered data
Global Const BASS_SPLIT_POS = &H2000

' Splitter attributes
Global Const BASS_ATTRIB_SPLIT_ASYNCBUFFER = &H15010
Global Const BASS_ATTRIB_SPLIT_ASYNCPERIOD = &H15011

' Envelope node
Type BASS_MIXER_NODE
        pos As Long
        poshi As Long
        value As Single
End Type

' Envelope types
Global Const BASS_MIXER_ENV_FREQ = 1
Global Const BASS_MIXER_ENV_VOL = 2
Global Const BASS_MIXER_ENV_PAN = 3
Global Const BASS_MIXER_ENV_LOOP = &H10000 ' flag: loop
Global Const BASS_MIXER_ENV_REMOVE = &H20000 ' flag: remove at end

' Additional sync types
Global Const BASS_SYNC_MIXER_ENVELOPE = &H10200
Global Const BASS_SYNC_MIXER_ENVELOPE_NODE = &H10201

' Additional BASS_Mixer_ChannelSetPosition flag
Global Const BASS_POS_MIXER_RESET = &H10000 ' flag: clear mixer's playback buffer

' Additional BASS_Mixer_ChannelGetPosition mode
Global Const BASS_POS_MIXER_DELAY = 5

' BASS_CHANNELINFO types
Global Const BASS_CTYPE_STREAM_MIXER = &H10800
Global Const BASS_CTYPE_STREAM_SPLIT = &H10801

Declare Function BASS_Mixer_GetVersion Lib "bassmix.dll" () As Long

Declare Function BASS_Mixer_StreamCreate Lib "bassmix.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long) As Long
Declare Function BASS_Mixer_StreamAddChannel Lib "bassmix.dll" (ByVal handle As Long, ByVal channel As Long, ByVal flags As Long) As Long
Declare Function BASS_Mixer_StreamAddChannelEx64 Lib "bassmix.dll" Alias "BASS_Mixer_StreamAddChannelEx" (ByVal handle As Long, ByVal channel As Long, ByVal flags As Long, ByVal start As Long, ByVal starthi As Long, ByVal length As Long, ByVal lengthhi As Long) As Long
Declare Function BASS_Mixer_StreamGetChannels Lib "bassmix.dll" (ByVal handle As Long, ByRef channels As Long, ByVal count As Long) As Long

Declare Function BASS_Mixer_ChannelGetMixer Lib "bassmix.dll" (ByVal handle As Long) As Long
Declare Function BASS_Mixer_ChannelIsActive Lib "bassmix.dll" (ByVal handle As Long) As Long
Declare Function BASS_Mixer_ChannelFlags Lib "bassmix.dll" (ByVal handle As Long, ByVal flags As Long, ByVal mask As Long) As Long
Declare Function BASS_Mixer_ChannelRemove Lib "bassmix.dll" (ByVal handle As Long) As Long
Declare Function BASS_Mixer_ChannelSetPosition64 Lib "bassmix.dll" Alias "BASS_Mixer_ChannelSetPosition" (ByVal handle As Long, ByVal pos As Long, ByVal poshi As Long, ByVal mode As Long) As Long
Declare Function BASS_Mixer_ChannelGetPosition Lib "bassmix.dll" (ByVal handle As Long, ByVal mode As Long) As Long
Declare Function BASS_Mixer_ChannelGetPositionEx Lib "bassmix.dll" (ByVal handle As Long, ByVal mode As Long, ByVal delay As Long) As Long
Declare Function BASS_Mixer_ChannelGetLevel Lib "bassmix.dll" (ByVal handle As Long) As Long
Declare Function BASS_Mixer_ChannelGetLevelEx Lib "bassmix.dll" (ByVal handle As Long, ByRef levels As Single, ByVal length As Single, ByVal flags As Long) As Long
Declare Function BASS_Mixer_ChannelGetData Lib "bassmix.dll" (ByVal handle As Long, ByRef buffer As Any, ByVal length As Long) As Long
Declare Function BASS_Mixer_ChannelSetSync64 Lib "bassmix.dll" Alias "BASS_Mixer_ChannelSetSync" (ByVal handle As Long, ByVal type_ As Long, ByVal param As Long, ByVal paramhi As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_Mixer_ChannelRemoveSync Lib "bassmix.dll" (ByVal handle As Long, ByVal sync As Long) As Long
Declare Function BASS_Mixer_ChannelSetMatrix Lib "bassmix.dll" (ByVal handle As Long, ByRef matrix As Single) As Long
Declare Function BASS_Mixer_ChannelSetMatrixEx Lib "bassmix.dll" (ByVal handle As Long, ByRef matrix As Single, ByVal time As Single) As Long
Declare Function BASS_Mixer_ChannelGetMatrix Lib "bassmix.dll" (ByVal handle As Long, ByRef matrix As Single) As Long
Declare Function BASS_Mixer_ChannelSetEnvelope Lib "bassmix.dll" (ByVal handle As Long, ByVal type_ As Long, ByRef nodes As BASS_MIXER_NODE, ByVal count As Long) As Long
Declare Function BASS_Mixer_ChannelSetEnvelopePos64 Lib "bassmix.dll" Alias "BASS_Mixer_ChannelSetEnvelopePos" (ByVal handle As Long, ByVal type_ As Long, ByVal pos As Long, ByVal poshi As Long) As Long
Declare Function BASS_Mixer_ChannelGetEnvelopePos Lib "bassmix.dll" (ByVal handle As Long, ByVal type_ As Long, ByRef value As Single) As Long

Declare Function BASS_Split_StreamCreate Lib "bassmix.dll" (ByVal channel As Long, ByVal flags As Long, ByRef chanmap As Long) As Long
Declare Function BASS_Split_StreamGetSource Lib "bassmix.dll" (ByVal handle As Long) As Long
Declare Function BASS_Split_StreamGetSplits Lib "bassmix.dll" (ByVal handle As Long, ByRef splits As Long, ByVal count As Long) As Long
Declare Function BASS_Split_StreamReset Lib "bassmix.dll" (ByVal handle As Long) As Long
Declare Function BASS_Split_StreamResetEx Lib "bassmix.dll" (ByVal handle As Long, ByVal offset As Long) As Long
Declare Function BASS_Split_StreamGetAvailable Lib "bassmix.dll" (ByVal handle As Long) As Long

' 32-bit wrappers for 64-bit BASS functions
Function BASS_Mixer_StreamAddChannelEx(ByVal handle As Long, ByVal channel As Long, ByVal flags As Long, ByVal start As Long, ByVal length As Long) As Long
BASS_Mixer_StreamAddChannelEx = BASS_Mixer_StreamAddChannelEx64(handle, channel, flags, start, 0, length, 0)
End Function

Function BASS_Mixer_ChannelSetPosition(ByVal handle As Long, ByVal pos As Long, ByVal mode As Long) As Long
BASS_Mixer_ChannelSetPosition = BASS_Mixer_ChannelSetPosition64(handle, pos, 0, mode)
End Function

Function BASS_Mixer_ChannelSetSync(ByVal handle As Long, ByVal type_ As Long, ByVal param As Long, ByVal proc As Long, ByVal user As Long) As Long
BASS_Mixer_ChannelSetSync = BASS_Mixer_ChannelSetSync64(handle, type_, param, 0, proc, user)
End Function

Function BASS_Mixer_ChannelSetEnvelopePos(ByVal handle As Long, ByVal type_ As Long, ByVal pos As Long) As Long
BASS_Mixer_ChannelSetEnvelopePos = BASS_Mixer_ChannelSetEnvelopePos64(handle, type_, pos, 0)
End Function
