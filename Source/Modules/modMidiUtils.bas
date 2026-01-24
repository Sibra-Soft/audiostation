Attribute VB_Name = "ModMidiUtils"
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Module:           ModMidiUtils
' Description:      Util functions for MIDI playback
'
' Date Changed:     25-10-2021
' Date Created:     04-10-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public gisEnd As Boolean
Public gisCurrentDoEvents As Boolean
Public gisCurrentQueue As Boolean
Public gisCurrentFF As Boolean
Public gmThreadPriorityApp As Integer

Public Type MidiFFTracking
    nDetectedMessageNumber As Long
End Type

Public Const MB_INTEGERUBOUND = 32767
Public Const MB_LONGUBOUND = &H7FFFFFFF
Public Const MB_LOWNIBBLE = &HF
Public Const MB_HIGHNIBBLE = &HF0
Public Const MB_LOWBYTE = &HFF
Public Const MB_HIGHBYTE = &HFF00
Public Const MB_DOEVENTSPOLLING = 10 ' release resources enough so <5% cpu usage

Public Const MB_DEVICEID = &H10 ' most common id used by output devices

Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function LockWindowUpdate Lib "user32" ( _
 ByVal hwndLock As Long) As Long


Public Function MidiNoteString2Display(chrBuffer As String) As String
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    '-------------------------------------------------------------------
    ' e.g. note buffer like Chr$(60)+Chr$(15)+Chr$(55) converts to mnemonic, "c4d#0g3"
    
    ' middle C, is mnemonic c4, value 60, center of keyboard using 8' stops
    ' max note range is mnemonic c-1 to g9, value 0 to 127
    ' common keyboard note range is mnemonic c2 to c7, value 36 to 96
    
    ' (see notes - mnemonic2value spreadsheet for more details)
    ' (see references in Roland sound module documentations)
    
    ' WARNING,
    ' Some non-Roland documention shows mnemonic value 60 as c5.
    ' The reason for the discrepancy is unclear.
    '-------------------------------------------------------------------

    Dim cText As String
    Dim mPosition As Integer
    Dim mNote As Integer
    Dim ctextscale As String
    Dim ctextoctave As String

    For mPosition = 1 To Len(chrBuffer)
        If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
        
        mNote = Asc(Mid$(chrBuffer, mPosition, 1))
        
        ctextoctave = Trim$(str$(Int((mNote - 0) / 12) - 1))
        ctextscale = ""
        Select Case (mNote Mod 12) ' 0-based scale, 12-steps
         Case 0: ctextscale = "c"
         Case 1: ctextscale = "c#"
         Case 2: ctextscale = "d"
         Case 3: ctextscale = "d#"
         Case 4: ctextscale = "e"
         Case 5: ctextscale = "f"
         Case 6: ctextscale = "f#"
         Case 7: ctextscale = "g"
         Case 8: ctextscale = "g#"
         Case 9: ctextscale = "a"
         Case 10: ctextscale = "a#"
         Case 11: ctextscale = "b"
         Case Else
            Err.Raise 1, , "PROGRAM ERROR 454, invalid note, " & Trim$(str$(mNote))
        End Select
        
        cText = cText & ctextscale & ctextoctave
        If mPosition = MB_INTEGERUBOUND Then Exit For ' (see 1.00.605)
    Next mPosition
        
    MidiNoteString2Display = cText

    Exit Function
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Function

Public Function RoundVB5(myexpression As Variant, Optional numdec As Integer) As Double
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    ' VB5 does not have the Round function ;-(
    Dim xx As Double, yy As Double
    
    If numdec = 0 Then
        xx = Int(myexpression)
        yy = myexpression - xx
        If yy >= 0.5 Then xx = xx + 1
    Else
        xx = myexpression * 10 ^ numdec
        yy = xx - Int(xx)
        If yy >= 0.5 Then xx = xx + 1
        xx = Int(xx) / (10 ^ numdec)
    End If
    RoundVB5 = xx
    Exit Function
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Function

Public Function AmbientUserMode() As Boolean
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    On Error Resume Next: Err.Clear ' prevent crash or halt
    Debug.Assert (0 / 0) ' crash by generating a division by zero error in debug mode
    If Err.Number <> 0 Then AmbientUserMode = True
    'If Ambient.UserMode() = True Then, does not work since debug version
    'If Ambient.UserMode() = False Then, does works since runtime version
    'Debug.Assert helps test conditions in debug mode to catch bugs early or verify integrity of data

    ' Alternative is to use App.EXEName or GetModuleFileName()
    ' (see msdn library, section, ID: Q177636
    ' HOWTO: Check If Program Is Running in the IDE or an EXE File)

    Exit Function
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Function

Public Function GetChecksum(ByVal cMessage As String) As String
    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
    
    Dim tempchksum As Integer
    Dim I As Integer
   
    tempchksum = 0
    For I = 1 To Len(cMessage)
        tempchksum = ((tempchksum + Asc(Mid$(cMessage, I, 1))) And &HFF)
        'tempchksum = (tempchksum + Asc(Mid$(cMessage, i, 1))) Mod 256 ' alternative
        'tempchksum = (tempchksum + Asc(Mid$(cMessage, i, 1))) Mod &H100 ' alternative
        If I = MB_INTEGERUBOUND Then Exit For
    Next I
    tempchksum = -tempchksum And 127
    If tempchksum = &H80 Then tempchksum = 0

    GetChecksum = Chr$(tempchksum)

    Exit Function
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
End Function

