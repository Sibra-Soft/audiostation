Attribute VB_Name = "MIDI_CONST"
Option Explicit

'
' State Constants
'
Global Const MIDISTATE_CLOSED = 0
Global Const MIDISTATE_OPEN = 1
Global Const MIDISTATE_STARTED = 2
Global Const MIDISTATE_STOPPED = 3
Global Const MIDISTATE_PAUSED = 4
'
' MidiOut actions
'
Global Const MIDIOUT_NONE = 0
Global Const MIDIOUT_OPEN = 1
Global Const MIDIOUT_CLOSE = 2
Global Const MIDIOUT_RESET = 3
Global Const MIDIOUT_START = 4
Global Const MIDIOUT_STOP = 5
Global Const MIDIOUT_QUEUE = 6
Global Const MIDIOUT_SEND = 7
Global Const MIDIOUT_TIMER = 8
Global Const MIDIOUT_PAUSE = 9
'
' MidiIn actions
'
Global Const MIDIIN_NONE = 0
Global Const MIDIIN_OPEN = 1
Global Const MIDIIN_CLOSE = 2
Global Const MIDIIN_RESET = 3
Global Const MIDIIN_START = 4
Global Const MIDIIN_STOP = 5
Global Const MIDIIN_REMOVE = 6
'
' MidiOut device types
'
Global Const MIDIOUT_PORT = 0
Global Const MIDIOUT_SQUARESYNTH = 1
Global Const MIDIOUT_FMSYNTH = 2
Global Const MIDIOUT_MIDIMAPPER = 3
'
' MidiFile actions
'
Global Const MIDIFILE_NONE = 0
Global Const MIDIFILE_OPEN = 1
Global Const MIDIFILE_CLOSE = 2
Global Const MIDIFILE_CREATE = 3
Global Const MIDIFILE_SAVE = 4
Global Const MIDIFILE_CLEAR = 5
Global Const MIDIFILE_INSERT_MESSAGE = 6
Global Const MIDIFILE_MODIFY_MESSAGE = 7
Global Const MIDIFILE_DELETE_MESSAGE = 8
Global Const MIDIFILE_INSERT_TRACK = 9
Global Const MIDIFILE_DELETE_TRACK = 10
Global Const MIDIFILE_SAVE_AS = 11

'
' Standard MIDI File Meta Event Constants
'
Global Const META = &HFF
Global Const META_SEQUENCE_NUMBER = &H0
Global Const META_TEXT = &H1
Global Const META_COPYRIGHT = &H2
Global Const META_NAME = &H3
Global Const META_INST_NAME = &H4
Global Const META_LYRIC = &H5
Global Const META_MARKER = &H6
Global Const META_CUE_POINT = &H7
Global Const META_CHAN_PREFIX = &H20
Global Const META_EOT = &H2F
Global Const META_TEMPO = &H51
Global Const META_SMPTE_OFFSET = &H54
Global Const META_TIME_SIG = &H58
Global Const META_KEY_SIG = &H59
Global Const META_SEQ_SPECIFIC = &H7F


' MIDI System Command Filter
Global Const FILTER_MTC = &HF1          'filter MTC
Global Const FILTER_SNGPTR = &HF2       'filter Song Position Pointer F2
Global Const FILTER_SNGSEL = &HF3       'filter Song Select F3
Global Const FILTER_F4 = &HF4           'filter undefined F4
Global Const FILTER_F5 = &HF5           'filter undefined F5
Global Const FILTER_TUNE = &HF6         'filter Tune Request F6
Global Const FILTER_CLOCK = &HF8        'filter MIDI Clock F8
Global Const FILTER_F9 = &HF9           'filter undefined F9
Global Const FILTER_START = &HFA        'filter MIDI Start FA
Global Const FILTER_CONT = &HFB         'filter MIDI Continue FB
Global Const FILTER_STOP = &HFC         'filter MIDI Stop FC
Global Const FILTER_FD = &HFD           'filter undefined FD
Global Const FILTER_ACTIVE_SENSE = &HFE 'filter Active Sensing FE
Global Const FILTER_RESET = &HFF        'filter System Reset FF

'
' MIDI status messages
'
Global Const note_uit = &H80
Global Const note_aan = &H90
Global Const POLY_KEY_PRESS = &HA0
Global Const CONTROLLER_CHANGE = &HB0
Global Const program_verander = &HC0
Global Const CHANNEL_PRESSURE = &HD0
Global Const pitch_bend = &HE0
Global Const SYSEX = &HF0
Global Const MTC_QFRAME = &HF1
Global Const EOX = &HF7
Global Const MIDI_CLOCK = &HF8
Global Const MIDI_START = &HFA
Global Const MIDI_CONTINUE = &HFB
Global Const MIDI_STOP = &HFC

'
' CONTROLLER_CHANGE, MIDI Controller Numbers Constants
'
Global Const MOD_WHEEL = 1
Global Const BREATH_CONTROLLER = 2
Global Const FOOT_CONTROLLER = 4
Global Const PORTAMENTO_TIME = 5
Global Const MAIN_VOLUME = 7
Global Const BALANCE = 8
Global Const PAN = 10
Global Const EXPRESS_CONTROLLER = 11
Global Const DAMPER_PEDAL = 64
Global Const PORTAMENTO = 65
Global Const SOSTENUTO = 66
Global Const SOFT_PEDAL = 67
Global Const HOLD_2 = 69
Global Const EXTERNAL_FX_DEPTH = 91
Global Const TREMELO_DEPTH = 92
Global Const CHORUS_DEPTH = 93
Global Const DETUNE_DEPTH = 94
Global Const PHASER_DEPTH = 95
Global Const DATA_INCREMENT = 96
Global Const DATA_DECREMENT = 97

'
' Constants to simplify MIDI
'
Global Const TOTAL_MIDI_CHANNELS = 16
