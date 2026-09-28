Attribute VB_Name = "modAdioMidi"
Option Explicit

'-------------------------------------------------------------------
' Constants used by Midi Input and Output objects, as of version 2.10.025
' Constants used by Midi File objects, as of version 1.20.011
'
' Acknowledgements:
' - original author: Mabry inc.
' - Web Link: www.mabry.com
'-------------------------------------------------------------------

'
' State Constants
'
Public Const MIDISTATE_CLOSED = 0
Public Const MIDISTATE_OPENED = 1 ' use <> MIDISTATE_CLOSED instead
Public Const MIDISTATE_STARTED = 2
Public Const MIDISTATE_STOPPED = 3
Public Const MIDISTATE_PAUSED = 4 ' use <> MIDISTATE_STARTED and QueueTimeCurrent <> 0
Public Const MIDISTATE_OPEN = 1 ' for backward compatibility
'
' MidiIn actions
'
Public Const MIDIIN_NONE = 0
Public Const MIDIIN_OPEN = 1
Public Const MIDIIN_CLOSE = 2
Public Const MIDIIN_RESET = 3
Public Const MIDIIN_START = 4
Public Const MIDIIN_STOP = 5
Public Const MIDIIN_REMOVE = 6
Public Const MIDIIN_QUEUE = 7 ' (see 1.00.515)
Public Const MIDIIN_SORT = 8 ' (see 1.00.515)
'
' MidiOut actions
'
Public Const MIDIOUT_NONE = 0
Public Const MIDIOUT_OPEN = 1
Public Const MIDIOUT_CLOSE = 2
Public Const MIDIOUT_RESET = 3
Public Const MIDIOUT_START = 4
Public Const MIDIOUT_STOP = 5
Public Const MIDIOUT_QUEUE = 6
Public Const MIDIOUT_SEND = 7
Public Const MIDIOUT_TIMER = 8
Public Const MIDIOUT_PAUSE = 9
Public Const MIDIOUT_REMOVE = 10 ' new in midiio version 2.00.302
Public Const MIDIOUT_READ = 11 ' new in midiio version 1.20.115
'
' Midi Message Status
'
Public Const MIDIMESSAGESTATE_NONE = 0     ' permanently disabled, no message (default when array cleared)
Public Const MIDIMESSAGESTATE_ENABLED = 1   ' enable message
Public Const MIDIMESSAGESTATE_DISABLED = 2 ' disabled message temporarily by user
Public Const MIDIMESSAGESTATE_DISABLED2 = 3 ' reserved
Public Const MIDIMESSAGESTATE_DISABLED3 = 4 ' reserved
'Public Const MIDIMESSAGESTATE_DISABLEDX ...  ' disabled message temporarily by user
' anything that is not enabled, will be considered disabled
'
' Midi Message Pointer
'
Public Const MIDIMP = 0 ' reserved
Public Const MIDIMP_MESSAGE = 1 ' in/out
Public Const MIDIMP_DATA1 = 2 ' in/out
Public Const MIDIMP_DATA2 = 3 ' in/out
Public Const MIDIMP_TIME = 4 ' in/out
Public Const MIDIMP_MESSAGETAG = 5 ' out
Public Const MIDIMP_MESSAGESTATE = 6 ' out
Public Const MIDIMP_TIMEOPEN = 10 ' in/out (readonly)
Public Const MIDIMP_DAYOPEN = 11 ' in/out (readonly)
Public Const MIDIMP_TIMETEMPO = 12 ' out (readonly)
Public Const MIDIMP_TIMESTARTOPEN = 13 ' out (readonly)
Public Const MIDIMP_TIMECURRENTOPEN = 14 ' out (readonly)
Public Const MIDIMP_TIMEACTUALOPEN = 15 ' out (readonly)
Public Const MIDIMP_TIMERELTO2 = 16 ' in (readonly)
Public Const MIDIMP_TIMERELTO3 = 17 ' in (readonly)
Public Const MIDIMP_TIMERELTO4 = 18 ' in (readonly)
Public Const MIDIMP_ERROR = 30 ' in (readonly)
Public Const MIDIMP_ERRORLATE = 31 ' in (readonly)
Public Const MIDIMP_ERRORLONG = 32 ' in (readonly)
Public Const MIDIMP_ERRORCOUNT = 33 ' in (readonly)
Public Const MIDIMP_UBOUND = 40 ' may increase in future versions
'
' MidiOut device types
' (WARNING: Not an accurate representation of devices
' returned by the DeviceType property because different
' depending on operating system and sound card.)
Public Const MIDIOUT_PORT = 0
Public Const MIDIOUT_SQUARESYNTH = 1
Public Const MIDIOUT_FMSYNTH = 2
Public Const MIDIOUT_MIDIMAPPER = 3
'Public Const MIDIOUT_NONE = ...
'Public Const MIDIOUT_WAVESYNTH = ...
'Public Const MIDIOUT_DEFAULT = ...
'
' SendNoteOff actions
'
Public Const MIDIOFF_ANO = 1
Public Const MIDIOFF_ASO = 2
Public Const MIDIOFF_EACH = 3
Public Const MIDIOFF_RECENT = 4
'
' MidiFile actions
'
Public Const MIDIFILE_NONE = 0
Public Const MIDIFILE_OPEN = 1
Public Const MIDIFILE_CLOSE = 2
Public Const MIDIFILE_CREATE = 3
Public Const MIDIFILE_SAVE = 4
Public Const MIDIFILE_CLEAR = 5
Public Const MIDIFILE_INSERT_MESSAGE = 6
Public Const MIDIFILE_MODIFY_MESSAGE = 7
Public Const MIDIFILE_DELETE_MESSAGE = 8
Public Const MIDIFILE_INSERT_TRACK = 9
Public Const MIDIFILE_DELETE_TRACK = 10
Public Const MIDIFILE_SAVE_AS = 11
'
' Standard MIDI File Meta Event Constants
'
Public Const META = &HFF
Public Const META_SEQUENCE_NUMBER = &H0
Public Const META_TEXT = &H1
Public Const META_COPYRIGHT = &H2
Public Const META_NAME = &H3
Public Const META_INST_NAME = &H4
Public Const META_LYRIC = &H5
Public Const META_MARKER = &H6
Public Const META_CUE_POINT = &H7
Public Const META_CHAN_PREFIX = &H20
Public Const META_EOT = &H2F
Public Const META_TEMPO = &H51
Public Const META_SMPTE_OFFSET = &H54
Public Const META_TIME_SIG = &H58
Public Const META_KEY_SIG = &H59
Public Const META_SEQ_SPECIFIC = &H7F
Public Const MAX_META_EVENT = &H7F
'
' Standard MIDI status messages
'
Public Const NOTE_OFF = &H80
Public Const NOTE_ON = &H90
Public Const POLY_KEY_PRESS = &HA0
Public Const CONTROLLER_CHANGE = &HB0
Public Const PROGRAM_CHANGE = &HC0
Public Const CHANNEL_PRESSURE = &HD0
Public Const PITCH_BEND = &HE0
Public Const SYSEX = &HF0
Public Const MTC_QFRAME = &HF1
Public Const MTC_SNGPTR = &HF2
Public Const MTC_SNGSEL = &HF3
Public Const MTC_TUNE = &HF6
Public Const EOX = &HF7
Public Const MIDI_CLOCK = &HF8
Public Const MIDI_START = &HFA
Public Const MIDI_CONTINUE = &HFB
Public Const MIDI_STOP = &HFC
Public Const ACTIVE_SENSE = &HFE
Public Const SYSTEM_RESET = &HFF
'
' Standard CONTROLLER_CHANGE, MIDI Controller Numbers Constants
'
Public Const MOD_WHEEL = 1
Public Const BREATH_CONTROLLER = 2
Public Const FOOT_CONTROLLER = 4
Public Const PORTAMENTO_TIME = 5
Public Const MAIN_VOLUME = 7
Public Const BALANCE = 8
Public Const PAN = 10
Public Const EXPRESS_CONTROLLER = 11
Public Const DAMPER_PEDAL = 64 ' also known as, sustain
Public Const PORTAMENTO = 65
Public Const SOSTENUTO = 66
Public Const SOFT_PEDAL = 67
Public Const HOLD_2 = 69
Public Const EXTERNAL_FX_DEPTH = 91
Public Const TREMELO_DEPTH = 92
Public Const CHORUS_DEPTH = 93
Public Const DETUNE_DEPTH = 94
Public Const PHASER_DEPTH = 95
Public Const DATA_INCREMENT = 96
Public Const DATA_DECREMENT = 97
'
' MIDI Filter Property Constants
'
Public Const FILTER_MTC = &HF1          'filter Frame, MTC_QFRAME
Public Const FILTER_SNGPTR = &HF2       'filter Song Position Pointer, MTC_SNGPTR
Public Const FILTER_SNGSEL = &HF3       'filter Song Select, MTC_SNGSEL
Public Const FILTER_F4 = &HF4           'filter undefined
Public Const FILTER_F5 = &HF5           'filter undefined
Public Const FILTER_TUNE = &HF6         'filter Tune Request, MTC_TUNE
Public Const FILTER_CLOCK = &HF8        'filter MIDI Clock, MIDI_CLOCK
Public Const FILTER_F9 = &HF9           'filter undefined
Public Const FILTER_START = &HFA        'filter MIDI Start, MIDI_START
Public Const FILTER_CONT = &HFB         'filter MIDI Continue, MIDI_CONT
Public Const FILTER_STOP = &HFC         'filter MIDI Stop, MIDI_STOP
Public Const FILTER_FD = &HFD           'filter undefined
Public Const FILTER_ACTIVE_SENSE = &HFE 'filter Active Sensing, ACTIVE_SENSE
Public Const FILTER_RESET = &HFF        'filter System Reset, SYSTEM_RESET
'
' Common constants
'
Public Const TOTAL_MIDI_CHANNELS = 16
Public Const MB_MSECPERDAY = 86400000 ' (24*60*60*1000)
Public Const MB_SECPERDAY = 86400 ' (24*60*60)
Public Const MB_VERY_OLD_TIME = -86400000 ' very old time to send fast ignoring message time
Public Const MB_LONGDAYUBOUNDLESS = 24 ' in days, less than 24.8, limit of long data type

Public Const MB_TESTERROREVENT = -1 ' optional to use with ErrorRaise()
