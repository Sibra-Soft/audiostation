VERSION 5.00
Object = "{0518EEBD-7F0E-4513-8491-A0221C9008A2}#2.1#0"; "midiio2k.ocx"
Object = "{4424C993-EABF-4A03-9BA9-369E0F07466E}#1.2#0"; "midifl2k.ocx"
Begin VB.UserControl AdioMidiPlayer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer_Playing_General 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   675
      Top             =   1890
   End
   Begin VB.HScrollBar HScrollPlayerTime 
      Height          =   285
      Left            =   405
      TabIndex        =   0
      Top             =   1350
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Timer Timer_Playing_Midi 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   135
      Top             =   1890
   End
   Begin MidiioLib.MIDIOutput MidiOutput 
      Left            =   2565
      Top             =   1080
      _Version        =   131073
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin MidifileLib.MIDIFile MidiFile 
      Left            =   2565
      Top             =   540
      _Version        =   65538
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      Filename        =   ""
   End
   Begin VB.Image Image_Main 
      Height          =   480
      Left            =   0
      Picture         =   "AdioMidiPlayer.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "AdioMidiPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'// Enums
Public Enum PlayMethod
    [VirtualMidiSync]
    [Direct]
End Enum

'// Private vars
Private Const MB_OPTIONOPENDEFAULT = 0
Private Const MB_STREAMNUMBER = 1
Private Const MB_STREAMEMPTY = 2
Private Const MB_HSCROLLTIMESCALEOFFSET = 1000&
Private Const MB_HSCROLLMESSAGESCALEOFFSET = 10&
Private Const MB_STREAMNAME_1 = "stream"
Private Const MB_STREAMNAME_FF = "FFstream"

Private Runtime As Long
Private CurrentElapsedTime As Long
Private maxt As Integer
Private TrackVis(255) As Integer
Private TrackOffset As Long
Private TimeExpectedMessage As Long
Private TimeExpectedMessageRelToTempo As Long
Private TimeExpectedMessageRelToOpen As Long
Private TimeActualMessageRelToOpen As Long
Private MainStreamNumber As Integer
Private MainStreamGroup() As Integer
Private MainStreamOption As Integer

'// Public vars
Public LoadedFile As String
Public LoadedFileType As enumAdioSupportedFileTypes
Public State As enumAdioPlayState
Public RepeatMode As enumAdioRepeatMode
Public NumberOfTracks As Integer

'// Events
Public Event Ready()
Public Event Playing()
Public Event StartPlay()
Public Event Paused()
Public Event Stopped()
Public Event MediaEnded()
Public Event NewMediaFile(file As String)
Public Event Error(Description As String, Code As Long)
Public Event MidiTrack(Name As String, TrackNr As Integer)
Public Event MidiTrackAudioLevelChange(TrackNr As Integer, level As Integer)
Private Function GetTrackName(track As Integer) As String
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown

Dim I As Long, bnk As Integer, map As Integer
Dim s1 As String

MidiFile.TrackNumber = track
bnk = 0: map = 0: TrackVis(track) = 1

For I = 1 To MidiFile.MessageCount ' 1-based scale
    MidiFile.MessageNumber = I ' 1-based scale
    '
    'Meta Event
    '
    If (MidiFile.Message = 255) And (MidiFile.Data1 = 3 Or MidiFile.Data1 = 1) Then
        If (MidiFile.MsgText = "") Then
            GetTrackName = "Track" & Str(track) & " (null)"
        Else
            If GetTrackName = "" Then
                GetTrackName = MidiFile.MsgText
            End If
        End If
    End If
    
    If (MidiFile.Message >= &HB0 And MidiFile.Data1 = &H0) Then
        bnk = MidiFile.Data2
    End If
    
    If (MidiFile.Message >= &HB0 And MidiFile.Data1 = &H20) Then
        map = MidiFile.Data2
    End If
    
    If (MidiFile.Message >= &HC0 And MidiFile.Message < &HD0) Then
        ' Use next line if desired :)
        s1 = "Channel " + Str$(MidiFile.Message - &HC0 + 1) _
        + " - Patch: " + Str$(1 + MidiFile.Data1) _
        + "    Bank/Map: " + Str$(bnk) + "/" + Trim$(Str$(map))
        If GetTrackName = "" Then
            GetTrackName = s1
        End If
        Exit Function
    End If
Next I

If GetTrackName = "" And MidiFile.Message <> 255 Then
    GetTrackName = "Channel " + Str$(1 + MidiFile.Message And &HF) + " - No Patch"
End If

If MidiFile.Message = 255 Then ' empty track
    TrackVis(track) = 0
End If

Exit Function
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Function
Private Sub DisplayTrackNames()
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown

Dim m As Integer
Dim t As Integer
Dim I As Integer

If MidiFile.NumberOfTracks = 1 Then
    TrackOffset = 1
Else
    TrackOffset = 2
End If
maxt = MidiFile.NumberOfTracks

If maxt > 32 Then maxt = 32

t = 200
If maxt > 16 Then t = 200

For t = 1 To maxt
    If (t >= 2) Or (MidiFile.NumberOfTracks = 1) Then
        RaiseEvent MidiTrack(Trim(GetTrackName(t)), t - 1)
    End If
Next t

Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
Public Function GetProperties() As mdlAdioProperties
Dim ReturnModel As New mdlAdioProperties

ReturnModel.DurationInSeconds = Runtime
ReturnModel.DurationString = Helpers.SecondsToTimeSerial(ReturnModel.DurationInSeconds, SmallTimeSerial)

ReturnModel.ElapsedInSeconds = CurrentElapsedTime
ReturnModel.ElapsedString = Helpers.SecondsToTimeSerial(ReturnModel.ElapsedInSeconds, SmallTimeSerial)

ReturnModel.RemainingInSeconds = Runtime - CurrentElapsedTime
ReturnModel.RemainingString = Helpers.SecondsToTimeSerial(ReturnModel.RemainingInSeconds, SmallTimeSerial)

Set GetProperties = ReturnModel
End Function
Private Sub StopStuckNote()
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
If MidiOutput.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close

MidiOutput.SendNoteOff (4) ' clear recent notes for stuck notes and sustain

Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
Private Sub OpenQueueStream(ByRef mStreamNumber As Integer, ByRef cStreamName As String, ByRef MIDIOutput1 As MidiOutput)
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown

Dim backupelement As Integer

Call MidiStackPushCommon(backupelement, MidiOutput)

If Len(Trim$(cStreamName)) = 0 Then Err.Raise 1, , "missing name"
If mStreamNumber = 0 Then
    MidiOutput.StreamNumber = mStreamNumber
    MidiOutput.StreamName = cStreamName
    
    If MidiOutput.StreamNumberTotal = MidiOutput.StreamNumberMax Then Err.Raise 1, , "too many streams"
    
    MidiOutput.ActionStream = MIDIOUT_OPEN
    
    mStreamNumber = MidiOutput.StreamNumber
End If

MidiOutput.StreamNumber = mStreamNumber

If MIDIOutput1.StateStreamEx(0) = MIDISTATE_CLOSED Then Err.Raise 1, , "" ' old stream is specified but not open properly ?
If MIDIOutput1.StreamName = "" Then Err.Raise 1, , ""

ExitSection:
    Call MidiStackPopCommon(backupelement, MIDIOutput1)
    
Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
Private Sub WaitSortStream(ByVal mStreamNumber As Integer)
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown

Dim mCount As Integer
Dim nMessageUBound As Long
Dim mStateSortStreamPercent As Long

Dim backupstreamnumber As Integer
backupstreamnumber = MidiOutput.StreamNumber ' alternative

MidiOutput.StreamNumber = mStreamNumber
nMessageUBound = MidiOutput.StreamMessageUBound
MidiOutput.StreamNumber = backupstreamnumber ' not needed anymore

mCount = 0
Do While MidiOutput.StateSortStreamEx(mStreamNumber) <> MIDISTATE_CLOSED
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Sleep MB_DOEVENTSPOLLING ' release resources enough so <5% cpu usage
    
    If Int(mCount / 10) = mCount / 10 Then ' interval = 10 * MB_DOEVENTSPOLLING
        mStateSortStreamPercent = MidiOutput.StateSortStreamPercentEx(mStreamNumber) ' may sort logarithmically slower

        mCount = 0 ' reset so not overflow
    End If
    mCount = mCount + 1
Loop

ExitSection:
Exit Sub

ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
Private Sub QueueSong_ByMidi1Track()
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown

Dim I As Long
Dim backupscreenmousepointer As Integer
Dim backupstreammessagenumbermax As Long
Dim mGroupNumber As Integer
Dim mStreamNumber As Integer
Dim nLo As Long
Dim nHi As Long
Dim isEmpty As Boolean

Dim mTrackPhysical As Integer
Dim mTrackLogical As Integer
Dim isTrackMute As Boolean
Dim m As Long
Dim nMessageCount As Long
Dim nMessageTotal As Long
Dim mR As Long
Dim mC As Long

Dim MIDIOutput1_MP(0 To MIDIMP_UBOUND) As Long ' always start from zero
Dim nMP As Long
Dim tempmessage As Integer
Dim tempdata1 As Integer
Dim tempdata2 As Integer
Dim temptime As Long
Dim tempmessagetag As Long
Dim tempmessagestate As Integer
Dim templogonly As Boolean

Dim mMsgFF81TempoCount As Integer
Dim mMsgFF88TPQCount As Integer
Dim mMsgFF81TempoCountMax As Integer
Dim mMsgFF88TPQCountMax As Integer
Dim arMsgFF81Tempo() As Long
Dim arMsgFF88TPQ() As Long
Const MB_DIMENSION1UBOUND = 3
Const MB_TICK = 1
Const MB_VALUE = 2
Const MB_TICKNEXT = 3

Dim backuptempo As Long
Dim backupticksperquarternote As Integer
Dim backupnumerator As Integer
Dim backupdenominator As Integer
Dim dTicksPerMillisecond As Double
Dim nTicksBetweenEvents As Long
Dim nTicksRemaining As Long
Dim nMillisecondsBetweenEvents As Long
Dim nStreamTimeCurrent As Long
Dim nStreamTicksCurrent As Long
Dim isTrackTicks As Boolean
Dim nStreamTimeStart As Long
Dim isGlobal As Boolean
Dim isMsgFF81TempoChange As Boolean
Dim isMsgFF88TPQChange As Boolean
Dim isSortOutOfOrder As Boolean

Dim nStartRelativeToStream As Long
Dim nCurrentRelativeToStream As Long
Dim dTimeDifferenceOld As Double
Dim dTimeDifference100 As Double
Dim dTimeDifferenceNew As Double
Dim dTempo As Double
Dim nTempoCurrent As Long
Dim nTempoPrevious As Long
Dim isProcessTempo As Boolean

If (MidiFile.FileName = "") Then GoTo ExitEnd

Dim backupelement As Integer
Call MidiStackPushCommon(backupelement, MidiOutput)

backupscreenmousepointer = Screen.MousePointer
Screen.MousePointer = 11

ReDim MainStreamGroup(MidiFile.NumberOfTracks + 1, 2) ' 1-based scale, plus master track
If UBound(MainStreamGroup, 1) = 0 Then Err.Raise 1, , "Missing stream number."
For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
    isEmpty = False
    If mGroupNumber <= MidiFile.NumberOfTracks Then
        MidiFile.TrackNumber = mGroupNumber
        If MidiFile.MessageCount = 0 Then
            ' empty track
            isEmpty = True
        End If
    Else
        ' master track used as is
    End If
    MainStreamGroup(mGroupNumber, MB_STREAMEMPTY) = isEmpty
    
    'If isEmpty = False Then ' skip if empty, but too complicated to handle later
    mStreamNumber = 0 ' new stream
    Call OpenQueueStream(mStreamNumber, MB_STREAMNAME_1, MidiOutput)
    MainStreamGroup(mGroupNumber, MB_STREAMNUMBER) = mStreamNumber
    MidiOutput.StreamNumber = mStreamNumber
    
    ' Clear any data if stream not new
    MidiOutput.ActionStream = MIDIOUT_RESET
    
    ' Total for reference.
    ' Incl. global track, master track and empty track.
    NumberOfTracks = Trim$(Str$(Val(NumberOfTracks) + 1))
Next mGroupNumber

' Get statistics
nMessageTotal = 0
backupstreammessagenumbermax = MidiOutput.StreamMessageNumberMax
For m = 1 To MidiFile.NumberOfTracks
    MidiFile.TrackNumber = m
    nMessageTotal = nMessageTotal + MidiFile.MessageCount
Next m

'Me.Caption = "MFPlayer Example - Loading - " & Trim$(Str$(Int(100 * nMessageCount / nMessageTotal))) & "%"

' Get global tags for reference.
' Assume in one track if midi file format 0.
' Assume in first track if midi file format 1. May be in others but not standard.
'{
    ' Get tempo info
    MidiFile.TrackNumber = 1
    MidiFile.MessageNumber = 0
    backuptempo = MidiFile.Tempo
    backupticksperquarternote = MidiFile.TicksPerQuarterNote
    backupnumerator = MidiFile.Numerator
    backupdenominator = 2 ^ MidiFile.Denominator
    If backuptempo = 0 Then backuptempo = 600000 ' assume 100 beats per minute (tempo/2 = beats*2)
    If backupticksperquarternote = 0 Then backupticksperquarternote = 480 ' assume 100 beats per minute
    If backupnumerator = 0 Then backupnumerator = 4 ' assume time signature 4/4
    If backupdenominator = 0 Then backupdenominator = 4
    dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000#
    
    ReDim arMsgFF81Tempo(MB_DIMENSION1UBOUND, 0 To 1000)
    ReDim arMsgFF88TPQ(MB_DIMENSION1UBOUND, 0 To 1000)
    
    mMsgFF81TempoCount = 0
    arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) = 0 ' tick zero
    arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount) = backuptempo
    arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) = 0 ' not yet
    
    mMsgFF88TPQCount = 0
    arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) = 0 ' tick zero
    arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount) = backupticksperquarternote
    arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) = 0 ' not yet
    
    nMessageCount = 0
    isSortOutOfOrder = False
    For mTrackPhysical = 1 To MidiFile.NumberOfTracks ' 1-based scale
        MidiFile.TrackNumber = mTrackPhysical ' 1-based scale (first is global, second is track one)
        mTrackLogical = mTrackPhysical - 1 ' 0-based scale (zero is global, first is track one)
    
        nStreamTicksCurrent = 0
        For m = 1 To MidiFile.MessageCount ' 1-based scale
            If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
            
            MidiFile.MessageNumber = m ' 1-based scale
            
            nTicksBetweenEvents = MidiFile.time
            nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
            
            tempmessage = MidiFile.Message
            tempdata1 = MidiFile.Data1
            If tempmessage <> META Then 'ignore
            ElseIf tempdata1 = META_TEMPO Then ' tempo
                mMsgFF81TempoCount = mMsgFF81TempoCount + 1
                If mMsgFF81TempoCount > UBound(arMsgFF81Tempo, 2) Then _
                 ReDim Preserve arMsgFF81Tempo(MB_DIMENSION1UBOUND, UBound(arMsgFF81Tempo, 2) + 100) ' more space
                
                arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) = nStreamTicksCurrent
                arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount) = MidiFile.Tempo
                arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) = 0 ' not yet
                arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount - 1) = nStreamTicksCurrent ' save for reference
            
                If arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) < arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount - 1) Then _
                 isSortOutOfOrder = True ' when message not in first track
            
            ElseIf tempdata1 = 88 Then ' time sig
                mMsgFF88TPQCount = mMsgFF88TPQCount + 1
                If mMsgFF88TPQCount > UBound(arMsgFF88TPQ, 2) Then _
                 ReDim Preserve arMsgFF88TPQ(MB_DIMENSION1UBOUND, UBound(arMsgFF88TPQ, 2) + 100) ' more space
                
                arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) = nStreamTicksCurrent
                arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount) = MidiFile.TicksPerQuarterNote
                arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) = 0 ' not yet
                arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount - 1) = nStreamTicksCurrent ' save for reference
            
                If arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) < arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount - 1) Then _
                 isSortOutOfOrder = True ' when message not in first track
            
            End If
        
            nMessageCount = nMessageCount + 1
            
            If nMessageCount = backupstreammessagenumbermax Then Exit For ' reached limit
        Next m
        
        If nMessageCount = backupstreammessagenumbermax Then Exit For ' reached limit
    Next mTrackPhysical
    
    mMsgFF81TempoCountMax = mMsgFF81TempoCount
    mMsgFF88TPQCountMax = mMsgFF88TPQCount

    If isSortOutOfOrder = True Then
        SortArray arMsgFF81Tempo _
         , mDimensionToSort:=2, mPositionToSort:=MB_TICK, mSecondaryPositionToSort:=0 _
         , nLo:=LBound(arMsgFF81Tempo, 2), nHi:=mMsgFF81TempoCountMax, isSortValue:=True _
         , mDimensionX:=2
    
        SortArray arMsgFF88TPQ _
         , mDimensionToSort:=2, mPositionToSort:=MB_TICK, mSecondaryPositionToSort:=0 _
         , nLo:=LBound(arMsgFF88TPQ, 2), nHi:=mMsgFF88TPQCountMax, isSortValue:=True _
         , mDimensionX:=2
    End If

nMessageCount = 0
For mTrackPhysical = 1 To MidiFile.NumberOfTracks ' 1-based scale
    MidiFile.TrackNumber = mTrackPhysical ' 1-based scale (first is global, second is track one)
    mTrackLogical = mTrackPhysical - 1 ' 0-based scale (zero is global, first is track one)
    
    mGroupNumber = mTrackPhysical ' 1-based scale (zero is none, first is global, second is track one, last is master track)
    MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
    
    isTrackMute = False
    
    If isTrackMute = False Then
        nStreamTicksCurrent = 0
        nStreamTimeCurrent = 0
        nStreamTimeStart = 0
        nTempoPrevious = backuptempo
        mMsgFF81TempoCount = 0
        mMsgFF88TPQCount = 0
        For m = 1 To MidiFile.MessageCount ' 1-based scale
            If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown

            MidiFile.MessageNumber = m ' 1-based scale
        
            ' Get next time
            nTicksRemaining = MidiFile.time
            
            ' Insert global messages, if any
            ' if occurs before next message.
            Do
                ' Assume not scan entire array for tempo since sequential and sorted.
                isGlobal = False
                isMsgFF81TempoChange = False
                isMsgFF88TPQChange = False
                If mMsgFF81TempoCount <> mMsgFF81TempoCountMax _
                 And nStreamTicksCurrent + nTicksRemaining >= arMsgFF81Tempo(MB_TICKNEXT, mMsgFF81TempoCount) Then
                    isGlobal = True
                    isMsgFF81TempoChange = True
                    mMsgFF81TempoCount = mMsgFF81TempoCount + 1
                    nTicksBetweenEvents = arMsgFF81Tempo(MB_TICK, mMsgFF81TempoCount) - nStreamTicksCurrent
                
                ElseIf mMsgFF88TPQCount <> mMsgFF88TPQCountMax _
                 And nStreamTicksCurrent + nTicksRemaining >= arMsgFF88TPQ(MB_TICKNEXT, mMsgFF88TPQCount) Then
                    isGlobal = True
                    isMsgFF88TPQChange = True
                    mMsgFF88TPQCount = mMsgFF88TPQCount + 1
                    nTicksBetweenEvents = arMsgFF88TPQ(MB_TICK, mMsgFF88TPQCount) - nStreamTicksCurrent
                End If
            
                If isGlobal = False Then Exit Do ' none
            
                ' Get time
                ' Assuming all previous ticks were for one tempo only.
                ' Assuming shifting start time already compensated for.
                ' Assume tracking ticks is more accurate than tracking time.
                nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
                nStreamTimeCurrent = Round(nStreamTicksCurrent / dTicksPerMillisecond, 0) ' ticks to time
            
                ' Adjust for changes in tempo.
                If isMsgFF81TempoChange = True Then
                    ' New start time.
                    ' Shift start time to compensate based on
                    ' current time and tempo rate.
                    '{
                        ' estimated message current time
                        nStartRelativeToStream = 0
                        nCurrentRelativeToStream = nStreamTimeCurrent
                        dTimeDifferenceOld = nCurrentRelativeToStream - nStartRelativeToStream
                                
                        ' get estimated current message time back to 100% tempo
                        ' already determined
                        dTempo = CDbl(nTempoPrevious) / 600000# * 100# ' percent
                        dTimeDifference100 = dTimeDifferenceOld _
                         * (1# / (dTempo / 100#))
                        ' 1/x from other to 100%
                        
                        ' get estimated starting time of new stream
                        nTempoCurrent = arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount)
                        dTempo = CDbl(nTempoCurrent) / 600000# * 100# ' percent
                        dTimeDifferenceNew = dTimeDifference100 _
                         * (1# * (dTempo / 100#))
                        ' 1*x from 100% to other
                        nStartRelativeToStream = nCurrentRelativeToStream - dTimeDifferenceNew
            
                        nStreamTimeStart = nStreamTimeStart + nStartRelativeToStream
                        nTempoPrevious = nTempoCurrent
                    '}
                    
                    ' New tempo.
                    dTicksPerMillisecond = (CDbl(arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount)) / CDbl(nTempoCurrent)) * 1000#
                    'dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000# ' not applicable
                
                ElseIf isMsgFF88TPQChange = True Then
                    ' New time signature.
                    ' Change tick scale but not speed of music.
                    nTempoCurrent = arMsgFF81Tempo(MB_VALUE, mMsgFF81TempoCount)
                    dTicksPerMillisecond = (CDbl(arMsgFF88TPQ(MB_VALUE, mMsgFF88TPQCount)) / CDbl(nTempoCurrent)) * 1000#
                    'dTicksPerMillisecond = (CDbl(backupticksperquarternote) / CDbl(backuptempo)) * 1000# ' not applicable
                End If
            
                nTicksRemaining = nTicksRemaining - nTicksBetweenEvents
            
                'Exit Do
            Loop
            
            ' Get next message
            ' store in variables for speed
            tempmessagestate = MIDIMESSAGESTATE_ENABLED
            templogonly = False
            tempmessage = MidiFile.Message
            tempdata1 = MidiFile.Data1
            tempdata2 = MidiFile.Data2
            
            ' Tag notes to play on keyboard and VU meters
            tempmessagetag = 0
            If (tempmessage And &HF0) = NOTE_ON And tempdata2 <> 0 Then
                tempmessagetag = tempdata2 + 1& + (mTrackLogical * 1000&)
            End If
            
            ' Get next time
            ' Assuming all previous ticks were for one tempo only.
            ' Assuming shifting start time already compensated for different tempos.
            nTicksBetweenEvents = nTicksRemaining
            nStreamTicksCurrent = nStreamTicksCurrent + nTicksBetweenEvents ' always ticks, no rounding
            nStreamTimeCurrent = Round(nStreamTicksCurrent / dTicksPerMillisecond, 0) ' ticks to time
            temptime = nStreamTimeStart + nStreamTimeCurrent

            ' Get buffer (no temporary variable for speed)
            If tempmessage = SYSEX Then ' SYSEX message
                MidiOutput.buffer = Chr(SYSEX) & MidiFile.buffer
            End If

            ' Queue with MessagePointer
            MIDIOutput1_MP(MIDIMP_MESSAGESTATE) = tempmessagestate
            MIDIOutput1_MP(MIDIMP_MESSAGE) = tempmessage
            MIDIOutput1_MP(MIDIMP_DATA1) = tempdata1
            MIDIOutput1_MP(MIDIMP_DATA2) = tempdata2
            MIDIOutput1_MP(MIDIMP_TIME) = temptime
            MIDIOutput1_MP(MIDIMP_MESSAGETAG) = tempmessagetag
            MidiOutput.MessagePointer(MIDIOutput1_MP(0), UBound(MIDIOutput1_MP)) = 0

            MidiOutput.MessageLogOnly = templogonly

            ' Add to output queue
            MidiOutput.StreamMessageNumber = 0 ' append
            MidiOutput.ActionStream = MIDIOUT_QUEUE
            nMessageCount = nMessageCount + 1
            
            If MidiOutput.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' reached limit
        Next m
    End If ' isTrackMute

    If MidiOutput.StreamMessageNumber = backupstreammessagenumbermax Then Exit For ' reached limit
Next mTrackPhysical

mGroupNumber = UBound(MainStreamGroup, 1) ' last is master track
MainStreamGroup(mGroupNumber, MB_STREAMNUMBER) = mGroupNumber
MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)

' First message to describe stream for reference (optional)
MidiOutput.MessageState = MIDIMESSAGESTATE_ENABLED
MidiOutput.MessageLogOnly = True
MidiOutput.Message = META
MidiOutput.Data1 = META_MARKER ' pass type of marker (0 to 255)
MidiOutput.Data2 = 0 ' pass information (optional)
MidiOutput.buffer = "built at, " & time
MidiOutput.time = 0
MidiOutput.MessageTag = 0
MidiOutput.StreamMessageNumber = 0 ' append
MidiOutput.ActionStream = MIDIOUT_QUEUE

nLo = 0
nHi = 0

For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
    MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
    
    ' Find common range in time.
    If MidiOutput.StreamMessageLastTime(1) > nHi Then _
     nHi = MidiOutput.StreamMessageLastTime(1)
Next mGroupNumber

For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
    MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
    MidiOutput.MessageState = MIDIMESSAGESTATE_ENABLED
    MidiOutput.MessageLogOnly = True
    MidiOutput.Message = META
    MidiOutput.Data1 = META_MARKER ' pass type of marker (0 to 255)
    MidiOutput.Data2 = 0 ' pass information (optional)
    MidiOutput.buffer = "Lowest time"
    MidiOutput.time = nLo
    MidiOutput.MessageTag = 0
    MidiOutput.StreamMessageNumber = 0 ' append
    MidiOutput.ActionStream = MIDIOUT_QUEUE
    MidiOutput.buffer = "Highest time"
    MidiOutput.time = nHi
    MidiOutput.StreamMessageNumber = 0 ' append
    MidiOutput.ActionStream = MIDIOUT_QUEUE
Next mGroupNumber

''''''''''''''''''''''''''''''''''''''''''''''
' Sort stream
''''''''''''''''''''''''''''''''''''''''''''''
For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
    MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
    
    If MidiOutput.StreamMessageSortOutOfOrder = False Then ' already queuesort
    ElseIf MidiOutput.StateStreamEx(0) = MIDISTATE_STARTED Then ' can only autosort
        Debug.Print "PROGRAM WARNING 21095, autosort"
    Else ' manualsort
        Call MidiOutput.SortStreamEx(MidiOutput.StreamNumber, 1) ' modeless
        Call WaitSortStream(MidiOutput.StreamNumber)
    End If
    
    If MidiOutput.StreamMessageSortOutOfOrder = True Then ' should have been sorted
        Err.Raise 1, , "PROGRAM ERROR 3276"
    End If
Next mGroupNumber

HScrollPlayerTime.max = nHi / 1000
Runtime = Round(nHi / 1000)

ExitSection:
    Screen.MousePointer = 0 ' backupscreenmousepointer
    Call MidiStackPopCommon(backupelement, MidiOutput)
    
Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
Private Sub PauseMidiPlayback()
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
If MidiOutput.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close

Dim mGroupNumber As Integer
Dim isPause As Boolean
Dim nTime As Long
Dim backupmessageventpause As Boolean

' Preserve passed data so not interfere with other functions
Dim backupelement As Integer
Call MidiStackPushCommon(backupelement, MidiOutput)

If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
    ' Midi format 0
    If MainStreamNumber <> 0 Then
        MidiOutput.StreamNumber = MainStreamNumber
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        ElseIf MidiOutput.StateStreamEx(0) = MIDISTATE_STARTED Then ' started
            MidiOutput.ActionStream = MIDIOUT_PAUSE
            Call StopStuckNote
        
        ElseIf MidiOutput.StateStreamEx(0) = MIDISTATE_STOPPED And MidiOutput.StreamTimeCurrent = 0 Then
            ' Start from stop with same pause button
            ' (not practical)
        
        ElseIf MidiOutput.StateStreamEx(0) = MIDISTATE_STOPPED And MidiOutput.StreamTimeCurrent > 0 Then
            MidiOutput.FilterLateEventStreamMax = True ' may filter notes
            MidiOutput.ActionStream = MIDIOUT_START
        End If
    End If

Else
    ' Midi format 1
    isPause = False
    
    ' Pause all streams at exact same time.
    backupmessageventpause = MidiOutput.MessageEventPause
    If MidiOutput.State <> MIDISTATE_CLOSED Then _
     MidiOutput.MessageEventPause = True
    
    For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
        MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        ElseIf MidiOutput.StateStreamEx(0) = MIDISTATE_STARTED Then ' started
            ' Pause from start
            MidiOutput.ActionStream = MIDIOUT_PAUSE
            Call StopStuckNote
            isPause = True
        
        ElseIf MidiOutput.StateStreamEx(0) = MIDISTATE_STOPPED And MidiOutput.StreamTimeCurrent = 0 Then
            ' Start from stop with same pause button
            ' (not practical)
        
        ElseIf MidiOutput.StateStreamEx(0) = MIDISTATE_STOPPED And MidiOutput.StreamTimeCurrent > 0 Then

            MidiOutput.FilterLateEventStreamMax = True ' may filter notes
            
            If nTime = 0 Then ' get once, same for all streams

            End If
            
            MidiOutput.ActionStream = MIDIOUT_START
        End If
    Next mGroupNumber
    
    If isPause = True Then Call StopStuckNote ' do again in case some streams were still processing
    If MidiOutput.State <> MIDISTATE_CLOSED Then MidiOutput.MessageEventPause = backupmessageventpause ' restore
End If

ExitSection:
    Call MidiStackPopCommon(backupelement, MidiOutput)

Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
Private Sub StopMidiPlayback()
If Not State = AdioPlaying Then: Exit Sub
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
If MidiOutput.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close

Dim mGroupNumber As Integer
Dim isStop As Boolean
Dim backupelement As Integer

Call MidiStackPushCommon(backupelement, MidiOutput)

If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
    If MainStreamNumber <> 0 Then
        MidiOutput.StreamNumber = MainStreamNumber
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        Else
            MidiOutput.ActionStream = MIDIOUT_STOP
            Call StopStuckNote
            'Call ClearScrollBar
        End If
    End If
Else
    isStop = False
    For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
        MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        Else
            MidiOutput.ActionStream = MIDIOUT_STOP
            Call StopStuckNote
            isStop = True
        End If
    Next mGroupNumber
    
    If isStop = True Then
        Call StopStuckNote ' do again in case some streams were still processing
    End If
End If

HScrollPlayerTime.Value = 0

ExitSection:
    Call MidiStackPopCommon(backupelement, MidiOutput)
    Exit Sub
ExitEnd: '
End Sub
Private Sub StartMidiPlayback()
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
If MidiOutput.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close

Dim mGroupNumber As Integer
Dim nTime As Long
Dim backupelement As Integer

Call MidiStackPushCommon(backupelement, MidiOutput)

If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
    If MainStreamNumber <> 0 Then
        MidiOutput.StreamNumber = MainStreamNumber
        
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        ElseIf MidiOutput.StateStreamEx(0) <> MIDISTATE_STARTED Then
            MidiOutput.FilterLateEventStreamMax = True ' may filter notes
            MidiOutput.StreamTimeStartRelativeToOpen = MidiOutput.TimeRelativeToOpen - HScrollPlayerTime.Value * MB_HSCROLLTIMESCALEOFFSET * (MidiOutput.StreamTempoRate / 100)
            MidiOutput.ActionStream = MIDIOUT_START
        End If
    End If

Else
    For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
        MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        ElseIf MidiOutput.StateStreamEx(0) <> MIDISTATE_STARTED Then
            MidiOutput.FilterLateEventStreamMax = True ' may filter notes

            If nTime = 0 Then ' get once, same for all streams
                nTime = MidiOutput.TimeRelativeToOpen - HScrollPlayerTime.Value * MB_HSCROLLTIMESCALEOFFSET * (MidiOutput.StreamTempoRate / 100)
            End If
            
            MidiOutput.StreamTimeStartRelativeToOpen = nTime
            MidiOutput.ActionStream = MIDIOUT_START
        End If
    Next mGroupNumber
End If

ExitSection:
    Call MidiStackPopCommon(backupelement, MidiOutput)
Exit Sub

ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
Public Sub StartPlay()
Dim IsMidi As Boolean

Select Case LoadedFileType

    Case 0, 1
        IsMidi = True
        Call StartMidiPlayback
    
    Case 2
        Call modPlaySid.StartPlay
    
    Case 3
        modPlayMus.Tempo = 120
        Call modPlayMus.StartPlay

End Select

State = AdioPlaying

RaiseEvent StartPlay
RaiseEvent Playing

If IsMidi Then
    Timer_Playing_Midi.Enabled = True
Else
    Timer_Playing_General.Enabled = True
End If
End Sub
Public Sub StopPlay()
Dim IsMidi As Boolean

Select Case LoadedFileType

    Case 0, 1
        IsMidi = True
        Call StopMidiPlayback
    
    Case 2
        Call modPlaySid.StopPlay
    
    Case 3
        Call modPlayMus.StopPlay

End Select

State = AdioStopped
RaiseEvent Stopped

If IsMidi Then
    Timer_Playing_Midi.Enabled = False
Else
    Timer_Playing_General.Enabled = False
End If
End Sub
Public Sub PausePlay()
Dim IsMidi As Boolean

Select Case LoadedFileType

    Case 0, 1
        IsMidi = True
        Call PauseMidiPlayback
        
    Case 2
        Call modPlaySid.PausePlay

End Select

State = AdioPaused
RaiseEvent Paused
End Sub
Public Sub SeekBySeconds(Direction As enumAdioSeekDirection, Optional seconds As Integer = 10)
Select Case Direction
    Case enumAdioSeekDirection.AdioForward
        HScrollPlayerTime.Value = HScrollPlayerTime.Value + seconds
    
    Case enumAdioSeekDirection.AdioRewind
        If seconds > HScrollPlayerTime.Value Then: Exit Sub
        HScrollPlayerTime.Value = HScrollPlayerTime.Value - seconds
End Select

Call ScrollBarPlayerTime_Forward0Common
End Sub
Public Sub InitComponent(DeviceId As Integer)
If MidiOutput.State <> MIDISTATE_CLOSED Then Call UnloadComponent

MidiOutput.DeviceId = DeviceId
MidiOutput.action = MIDIOUT_OPEN

If MidiOutput.ErrorCode = 0 Then: RaiseEvent Ready
State = AdioReady
End Sub
Private Sub UnloadComponent()
modPlaySid.StopPlay
modPlayMus.StopPlay

If MidiOutput.State <> MIDISTATE_CLOSED Then
    StopPlay ' clear any stuck notes
    MidiOutput.action = MIDIOUT_CLOSE
End If
End Sub
Public Function GetListOfMidiDevices() As Collection
Dim I As Integer
Dim MidiDevice As mdlAdioMidiDevice
Dim ListOfDevices As New Collection

For I = -1 To MidiOutput.DeviceCount - 1
    Set MidiDevice = New mdlAdioMidiDevice
    
    MidiOutput.DeviceId = I
    
    MidiDevice.mId = I + 1
    MidiDevice.mName = MidiOutput.ProductName
    
    ListOfDevices.Add MidiDevice
Next I

Set GetListOfMidiDevices = ListOfDevices
End Function

Public Sub LoadFile(strFile As String)
Dim Fso As New FileSystemObject

If Not Helpers.FileExists(strFile) Then: RaiseEvent Error("File not found", 100)
If Not CheckFileSupport(strFile) Then: RaiseEvent Error("File not supported", 110)

Select Case LCase(Fso.GetExtensionName(strFile))

    Case "mid", "midi"
        LoadedFileType = [MID - Midi File]
        Call LoadMidiFile(strFile)
        
    Case "kar"
        LoadedFileType = [KAR - Karaoke File]
        Call LoadMidiFile(strFile)
        
    Case "mus"
        LoadedFileType = [MUS - Beep Symfony File]
        Call modPlayMus.LoadFile(strFile)
    
    Case "sid"
        LoadedFileType = [SID - Commodore64 8Bit File]
        Call modPlaySid.LoadFile(strFile)
    
End Select

LoadedFile = strFile

RaiseEvent NewMediaFile(strFile)
End Sub
Private Sub LoadMidiFile(strFile As String)
MidiFile.action = MIDIFILE_CLOSE
MidiFile.ReadOnly = True
MidiFile.FileName = strFile
MidiFile.action = MIDIFILE_OPEN

Call DisplayTrackNames
Call QueueSong_ByMidi1Track

HScrollPlayerTime.Value = 0
MainStreamOption = 0
End Sub
Private Sub MidiOutput_StreamSend()
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown

Dim MessageTag(TOTAL_MIDI_CHANNELS) As Long
Dim TrackNumber As Integer
Dim Intensity As Integer
Dim Channel As Integer
Dim m As Integer
Dim nTime As Long
Dim nTimeRelToTempo As Long

'
' Get last messagetag in each channel
'
nTime = MidiOutput.SendTime ' initialize
nTimeRelToTempo = MidiOutput.SendTimeRelativeToTempo ' initialize
Do While MidiOutput.MessageCount > 0 ' (optional)
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Channel = (MidiOutput.SendMessage And &HF) + 1
    If MidiOutput.SendMessageTag = 0 Then ' no tag
    ElseIf MidiOutput.SendMessageTag < MessageTag(Channel) Then ' not a peak
    Else
        ' Overwrite and discard old message tags, if any
        MessageTag(Channel) = MidiOutput.SendMessageTag
    End If
    
    ' Track oldest message to reflect in scroll bar
    ' (optional)
    If MidiOutput.SendTime < nTime Then nTime = MidiOutput.SendTime
    If MidiOutput.SendTimeRelativeToTempo < nTimeRelToTempo Then nTimeRelToTempo = MidiOutput.SendTimeRelativeToTempo
    
    MidiOutput.ActionStream = MIDIOUT_REMOVE
Loop

TimeExpectedMessage = nTime
TimeExpectedMessageRelToTempo = nTimeRelToTempo

For Channel = 1 To TOTAL_MIDI_CHANNELS
    If MessageTag(Channel) = 0 Then ' no tag
    ElseIf (MessageTag(Channel) < 0) Or (MessageTag(Channel) >= 32000) Then ' not applicable
    Else
        Intensity = MessageTag(Channel) Mod 1000
        TrackNumber = Int(MessageTag(Channel) / 1000)
        TrackNumber = TrackNumber + 1 ' restore to 1-based scale to match other arrays
        
        RaiseEvent MidiTrackAudioLevelChange((TrackNumber - TrackOffset), Intensity)
    End If
Next Channel

Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub
Private Sub ScrollBarPlayerTime_Forward0Common()
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
If MidiOutput.State = MIDISTATE_CLOSED Then GoTo ExitEnd ' not needed at close

Dim isStarted As Boolean
Dim mGroupNumber As Integer
Dim nSeekFromTime As Long
Dim nSeekToTime As Long
Dim nSeekFromMessage As Long
Dim nSeekToMessage As Long

Dim backuplabel As String

isStarted = False

Dim backupelement As Integer
Call MidiStackPushCommon(backupelement, MidiOutput)

gisCurrentFF = True ' prevent multithreading issues caused by doevents

If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
    ' Midi format 0
    If MainStreamNumber <> 0 Then
        MidiOutput.StreamNumber = MainStreamNumber
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        ElseIf MidiOutput.StateStreamEx(0) <> MIDISTATE_STARTED Then ' not started
        Else
            isStarted = True
        End If

        MidiOutput.ActionStream = MIDIOUT_PAUSE
        Call StopStuckNote
        nSeekFromTime = MidiOutput.StreamTimeCurrent
        nSeekToTime = HScrollPlayerTime.Value * MB_HSCROLLTIMESCALEOFFSET

        If nSeekToTime >= nSeekFromTime Then

        Else
            nSeekFromTime = 0
            MidiOutput.StreamTimeCurrent = nSeekFromTime
        End If
    End If

Else
    ' Midi format 1
    For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
        ' Stop all streams quickly
        MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        ElseIf MidiOutput.StateStreamEx(0) <> MIDISTATE_STARTED Then ' not started
        Else
            isStarted = True
        End If

        MidiOutput.ActionStream = MIDIOUT_PAUSE
        Call StopStuckNote
    Next mGroupNumber
    
    For mGroupNumber = 1 To UBound(MainStreamGroup, 1)
        ' Process fast forward
        MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        nSeekFromTime = MidiOutput.StreamTimeCurrent
        nSeekToTime = HScrollPlayerTime.Value * MB_HSCROLLTIMESCALEOFFSET

        If nSeekToTime >= nSeekFromTime Then

        Else
            nSeekFromTime = 0
            MidiOutput.StreamTimeCurrent = nSeekFromTime
        End If
    Next mGroupNumber
End If

' Keep started since paused.
' Approximate time by shifting start time.
If isStarted = True Then StartPlay

ExitSection:
    gisCurrentFF = False ' not needed anymore
    Call MidiStackPopCommon(backupelement, MidiOutput)
    Exit Sub
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub Timer_Playing_General_Timer()
CurrentElapsedTime = CurrentElapsedTime + 1
End Sub

Private Sub Timer_Playing_Midi_Timer()
If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown

Dim mGroupNumber As Integer
Dim nTimeExpectedStream As Long
Dim nTime As Long
Dim backupelement As Integer

Call MidiStackPushCommon(backupelement, MidiOutput)

If MainStreamOption <> MB_OPTIONOPENDEFAULT Then
    ' Midi format 0
    If MainStreamNumber <> 0 Then
        MidiOutput.StreamNumber = MainStreamNumber
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        ElseIf HScrollPlayerTime.tag > Trim$(Str$(time - 2 / 86400)) Then ' still scrolling two sec
        Else

            CurrentElapsedTime = (MidiOutput.StreamTimeCurrent / 1000)
            HScrollPlayerTime.Value = CInt(MidiOutput.StreamTimeCurrent / 1000)
            
        End If
    End If

Else
    ' Midi format 1
    mGroupNumber = UBound(MainStreamGroup, 1) ' last is master track
    
    If mGroupNumber <> 0 Then
        MidiOutput.StreamNumber = MainStreamGroup(mGroupNumber, MB_STREAMNUMBER)
        If MidiOutput.StateStreamEx(0) = MIDISTATE_CLOSED Then ' no stream
        ElseIf HScrollPlayerTime.tag > Trim$(Str$(time - 2 / 86400)) Then ' still scrolling two sec
        Else
        
            CurrentElapsedTime = (MidiOutput.StreamTimeCurrent / 1000)
            HScrollPlayerTime.Value = CInt(MidiOutput.StreamTimeCurrent / 1000)
            
        End If
    End If
End If

If CurrentElapsedTime = Runtime Then
    State = AdioEnded
    
    RaiseEvent MediaEnded
    Timer_Playing_Midi.Enabled = False
End If

ExitSection:
    Call MidiStackPopCommon(backupelement, MidiOutput)
    Exit Sub

ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Sub

Private Sub UserControl_Initialize()
MidiOutput.ErrorScheme = 1
MidiOutput.ErrorHalt = True
End Sub

Private Sub UserControl_Resize()
width = Image_Main.width
height = Image_Main.height
End Sub

Private Sub UserControl_Terminate()
gisEnd = True
UnloadComponent
End Sub

