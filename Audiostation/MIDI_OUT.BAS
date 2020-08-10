Attribute VB_Name = "MIDI_OUT"
Option Explicit
Dim m_hmidiout As Long

Global Const MIDIERR_BASE = 64

'  general constants
Global Const MAXPNAMELEN = 32           '  max product name length (including NULL)
Global Const MAXERRORLENGTH = 128       '  max error text length (including NULL)


Global Const MM_MIM_OPEN = &H3C1                    '  MIDI input
Global Const MM_MIM_CLOSE = &H3C2
Global Const MM_MIM_DATA = &H3C3
Global Const MM_MIM_LONGDATA = &H3C4
Global Const MM_MIM_ERROR = &H3C5
Global Const MM_MIM_LONGERROR = &H3C6

Global Const MM_MOM_OPEN = &H3C7                    '  MIDI output
Global Const MM_MOM_CLOSE = &H3C8
Global Const MM_MOM_DONE = &H3C9

' ***************************************************************************

'                             MIDI audio support

' ****************************************************************************/

'  MIDI error return values
Global Const MIDIERR_UNPREPARED = (MIDIERR_BASE + 0)       '  header not prepared
Global Const MIDIERR_STILLPLAYING = (MIDIERR_BASE + 1)     '  still something playing
Global Const MIDIERR_NOMAP = (MIDIERR_BASE + 2)            '  no current map
Global Const MIDIERR_NOTREADY = (MIDIERR_BASE + 3)         '  hardware is still busy
Global Const MIDIERR_NODEVICE = (MIDIERR_BASE + 4)         '  port no longer connected
Global Const MIDIERR_INVALIDSETUP = (MIDIERR_BASE + 5)     '  invalid setup
Global Const MIDIERR_LASTERROR = (MIDIERR_BASE + 5)        '  last error in range

Global Const MIDIPATCHSIZE = 128

'  MIDI callback messages
Global Const MIM_OPEN = MM_MIM_OPEN
Global Const MIM_CLOSE = MM_MIM_CLOSE
Global Const MIM_DATA = MM_MIM_DATA
Global Const MIM_LONGDATA = MM_MIM_LONGDATA
Global Const MIM_ERROR = MM_MIM_ERROR
Global Const MIM_LONGERROR = MM_MIM_LONGERROR
Global Const MOM_OPEN = MM_MOM_OPEN
Global Const MOM_CLOSE = MM_MOM_CLOSE
Global Const MOM_DONE = MM_MOM_DONE

'  device ID for MIDI mapper
Global Const MIDIMAPPER = (-1)
Global Const MIDI_MAPPER = (-1)

'  flags for wFlags parm of midiOutCachePatches(), midiOutCacheDrumPatches()
Global Const MIDI_CACHE_ALL = 1
Global Const MIDI_CACHE_BESTFIT = 2
Global Const MIDI_CACHE_QUERY = 3
Global Const MIDI_UNCACHE = 4
'  flags for wTechnology field of MIDIOUTCAPS structure
Global Const MOD_MIDIPORT = 1      '  output port
Global Const MOD_SYNTH = 2         '  generic internal synth
Global Const MOD_SQSYNTH = 3       '  square wave internal synth
Global Const MOD_FMSYNTH = 4       '  FM internal synth
Global Const MOD_MAPPER = 5        '  MIDI mapper

'  flags for dwSupport field of MIDIOUTCAPS structure
Global Const MIDICAPS_VOLUME = &H1               '  supports volume control
Global Const MIDICAPS_LRVOLUME = &H2             '  separate left-right volume control
Global Const MIDICAPS_CACHE = &H4

'  MIDI output device capabilities structure
'Type MIDIINCAPS
'    wMid As Integer                  '  manufacturer ID
'    wPid As Integer                  '  product ID
'    vDriverVersion As Integer        '  version of the driver
'    szPname As String * MAXPNAMELEN  '  product name (NULL terminated string)
'End Type
'  flags for dwFlags field of MIDIHDR structure
Global Const MHDR_DONE = &H1                     '  done bit
Global Const MHDR_PREPARED = &H2                 '  set if header prepared
Global Const MHDR_INQUEUE = &H4                  '  reserved for driver
' MIDI output device capabilities structure
Type MIDIOUTCAPS
    wMid As Integer                ' Manufacturer ID
    wPid As Integer                ' Product ID
    vDriverVersion As Long         ' Driver version
    szPname As String * 32         ' Product name (NULL terminated string)
    wTechnology As Integer         ' Device type
    wVoices As Integer             ' n. of voices (internal synth only)
    wNotes As Integer              ' max n. of notes (internal synth only)
    wChannelMask As Integer        ' n. of Midi channels (internal synth only)
    dwSupport As Long              ' Supported extra controllers (volume, etc)
End Type
' MIDI data block header
Type MIDIHDR
    lpData As String             '  pointer to locked data block
    dwBufferLength As Long       '  length of data in data block
    dwBytesRecorded As Long      '  used for input only
    dwUser As Long               '  for client's use
    dwFlags As Long              '  assorted flags (see defines)
    lpNext As Long               '  reserved for driver
    Reserved As Long             '  reserved for driver
End Type
'
' 32 bit
'
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Private Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Private Declare Function MIDIOutOpen Lib "winmm.dll" Alias "midiOutOpen" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long


Sub midi_listoutdevs(c As Control)
Dim I As Integer
Dim X As Integer
Dim midicaps As MIDIOUTCAPS

    c.Clear
    ' Test for MIDI mapper
    If midiOutGetDevCaps(MIDIMAPPER, midicaps, Len(midicaps)) = 0 Then ' OK
        c.AddItem midicaps.szPname
        c.ItemData(c.NewIndex) = MIDIMAPPER ' Save dev_id in item data
    End If
    ' Add other devs
    For I = 0 To midiOutGetNumDevs() - 1
        If midiOutGetDevCaps(I, midicaps, Len(midicaps)) = 0 Then ' OK
            c.AddItem midicaps.szPname
            c.ItemData(c.NewIndex) = I ' Save dev_id
        End If
    Next
End Sub

Sub midi_out_close()
Dim midi_error As Integer

    If m_hmidiout <> 0 Then
        midi_error = midiOutClose(m_hmidiout)
        If Not midi_error = 0 Then
            Call midi_outerr(midi_error)
        End If
        m_hmidiout = 0
    End If
End Sub

Function midi_out_open(ByVal dev_id As Integer) As Integer
Dim midi_error As Integer

    midi_out_close ' just in case (And it dont hurt)
    midi_error = MIDIOutOpen(m_hmidiout, dev_id, 0, 0, 0)
    If Not midi_error = 0 Then
        Call midi_outerr(midi_error)
    End If
    midi_out_open = (m_hmidiout <> 0)
End Function

Sub midi_outerr(ByVal midi_error As Integer)
Dim s As String
Dim X As Integer

    s = Space(MAXERRORLENGTH)
    X = midiOutGetErrorText(midi_error, s, MAXERRORLENGTH)
    'If Not g_debug Then
        MsgBox s
    'End If
End Sub

Sub midi_outshort(b1 As Integer, b2 As Integer, b3 As Integer)
Dim midi_error As Integer

    midi_error = midiOutShortMsg(m_hmidiout, packdword(0, b3, b2, b1))
    If Not midi_error = 0 Then
        Call midi_outerr(midi_error)
    End If
End Sub

Function packdword(i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer) As Long
    packdword = i2 * &H10000 + i3 * &H100 + i4
End Function

