VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSibraSoft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////
'// FileName        : SibraSoft.cls
'// FileType        : Microsoft Visual Basic 6 - Class Module
'// Type            : Class
'// Author          : Alex van den Berg
'// Created         : 25-06-2022
'// Last Modified   : 03-11-2023
'// Copyright       : Sibra-Soft
'// Description     : Various functions and extensions used in all projects
'////////////////////////////////////////////////////////////////

Option Explicit

Public Enum enumTimeSerial
    [LongTimeSerial] = 0
    [SmallTimeSerial] = 1
End Enum

Public Enum enumMsgType
    [LogError]
    [LogWarning]
    [LogInfo]
    [logSuccessAudit]
    [logFailureAudit]
End Enum

Private Type BrowseInfo
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFilename As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'INI
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Function Exists(ByVal oCol As Collection, ByVal vKey As Variant) As Boolean
On Error Resume Next
oCol.Item vKey
Exists = (Err.Number = 0)
Err.Clear
End Function

Public Sub CreateFolderIfNotExists(FolderPath As String)
If Dir(FolderPath, vbDirectory) = vbNullString Then MkDir FolderPath
End Sub
Public Function SetBytes(bytes) As String
On Error GoTo ErrorHandler
If bytes >= 1073741824 Then
    SetBytes = format(bytes / 1024 / 1024 / 1024, "#0.00") _
         & " GB"
ElseIf bytes >= 1048576 Then
    SetBytes = format(bytes / 1024 / 1024, "#0.00") & " MB"
ElseIf bytes >= 1024 Then
    SetBytes = format(bytes / 1024, "#0.00") & " KB"
ElseIf bytes < 1024 Then
    SetBytes = Fix(bytes) & " Bytes"
End If
Exit Function

ErrorHandler:
SetBytes = "0 Bytes"
End Function
Public Sub TerminateProcessByPid(Pid As Long)
Shell "taskkill.exe /f /t /pid " & Pid, vbHide
End Sub
Public Function ShellAndWait(command As String)
Dim oShell As Object

Set oShell = CreateObject("WSCript.shell")

oShell.Run "cmd /C " & Chr(34) & App.path & Chr(34) & "\" & command, 0, True
End Function
Function IsOddNumber(ByVal iNum As Integer) As Boolean
    IsOddNumber = ((iNum \ 2) * 2 <> iNum)
End Function
Public Function FileExists(sFullPath As String) As Boolean
Dim myFSO As Object
Set myFSO = CreateObject("Scripting.FileSystemObject")

FileExists = myFSO.FileExists(sFullPath)
End Function
Public Function CommondialogFilesToList(Files As String) As String
Dim FileList
Dim I As Integer
Dim Output As String

Debug.Print Files
FileList = Split(Files, Chr(0))

If UBound(FileList) = 0 Then
    Output = Files
Else
    For I = 0 To UBound(FileList)
        If Me.FileExists(FileList(0) & "\" & FileList(I)) Then
            Output = Output & FileList(0) & "\" & FileList(I) & vbNewLine
        End If
    Next
End If

CommondialogFilesToList = Output
End Function
Public Function RegistryRead(RegistryPath As String, RegistryKey As String) As Variant
Dim Registry As Object

Set Registry = CreateObject("WScript.Shell")
RegistryRead = Registry.RegRead(RegistryPath & RegistryKey)
End Function
Public Sub FilePutContents(File As String, Contents As String)
Dim iFileNo As Integer

iFileNo = FreeFile

Open File For Output As #iFileNo
    Print #iFileNo, Contents
Close #iFileNo
End Sub
Public Function FileGetContents(File As String) As String
Dim intFile As Integer
Dim strData As String

intFile = FreeFile

Open File For Input As #intFile
    FileGetContents = Input(LOF(intFile), #intFile)
Close #intFile
End Function
Public Function GetTimeRemaining(StartTime As Date, ItemsProcessed As Integer, TotalItems As Integer)
Dim ElapsedTime As Date
Dim EstimatedTotalTime As Date
Dim EstimatedTimeRemaining As Date
Dim TimePerIteration As Double

ElapsedTime = (Now - StartTime)
TimePerIteration = ElapsedTime / ItemsProcessed
EstimatedTotalTime = CDate((TotalItems * TimePerIteration))
EstimatedTimeRemaining = EstimatedTotalTime - ElapsedTime

GetTimeRemaining = EstimatedTimeRemaining
End Function
Public Function WriteEntryToWindowsEventLog(ErrorCode As Integer, MsgDescription As String, msgType As enumMsgType, msgLog As String)
Dim strErrorType As String

If msgType = LogError Then: strErrorType = "ERROR"
If msgType = LogWarning Then: strErrorType = "WARNING"
If msgType = LogInfo Then: strErrorType = "INFORMATION"
If msgType = logSuccessAudit Then: strErrorType = "SUCCESSAUDIT"
If msgType = logFailureAudit Then: strErrorType = "FAILUREAUDIT"

Shell ("cmd.exe /c EVENTCREATE /T " & strErrorType & " /ID " & ErrorCode & " /L " & msgLog & " /D " & Chr(34) & MsgDescription & Chr(34) & " /so fonvex.exe"), vbHide
End Function
Public Function FolderName(FileName As String) As String
Dim posn As Integer

posn = InStrRev(FileName, "\")

If posn > 0 Then
    FolderName = Left$(FileName, posn)
Else
    FolderName = ""
End If
End Function
Public Function MilliSecondsToSeconds(MillTime As Long) As Long
MilliSecondsToSeconds = (MillTime \ 1000) Mod 60
End Function
Public Function SecondsToTimeSerial(ByVal sec As Long, ReturnType As enumTimeSerial) As String
Dim mHours As Long, mMinutes As Long, mSeconds As Long

mSeconds = sec
mHours = mSeconds \ 3600
mMinutes = (mSeconds - (mHours * 3600)) \ 60
mSeconds = mSeconds - ((mHours * 3600) + (mMinutes * 60))

If ReturnType = LongTimeSerial Then
    SecondsToTimeSerial = format(mHours, "00") & ":" & format(mMinutes, "00") & ":" & format(mSeconds, "00")
Else
    SecondsToTimeSerial = format(mMinutes, "00") & ":" & format(mSeconds, "00")
End If
End Function
Public Function MilliSecondsToTimeSerial(ByVal MilliSeconds As Long, ReturnType As enumTimeSerial) As String
Dim lngSecRemainder As Long
Dim lngMinSecRemainder As Long
Dim lngHoursPart As Long
Dim lngMinutesPart As Long
Dim lngSecondsPart As Long
Dim sTimeRemaining As String
Dim sHoursPart As String
Dim sMinutesPart As String
Dim sSecondsPart As String
  
lngHoursPart = MilliSeconds \ 3600000
lngMinSecRemainder = MilliSeconds Mod 3600000
lngMinutesPart = lngMinSecRemainder \ 60000
lngSecRemainder = lngMinSecRemainder Mod 60000
lngSecondsPart = lngSecRemainder \ 1000

sHoursPart = format(lngHoursPart, "00")
sMinutesPart = format(lngMinutesPart, "00")
sSecondsPart = format(lngSecondsPart, "00")

If ReturnType = LongTimeSerial Then
    sTimeRemaining = sHoursPart & ":" & sMinutesPart & ":" & sSecondsPart
Else
    sTimeRemaining = sMinutesPart & ":" & sSecondsPart
End If

MilliSecondsToTimeSerial = sTimeRemaining
End Function


Public Function CollectionToString(TargetCollection As Collection, Seperator As String) As String
Dim I As Integer
Dim ReturnValue As String

For I = 1 To TargetCollection.Count
    ReturnValue = ReturnValue & TargetCollection(I) & Seperator
Next

CollectionToString = Left$(ReturnValue, Len(ReturnValue) - 1)
End Function
Public Function INIRead(section As String, KeyName As String, FileName As String, Optional DefaultValue As String) As String
Dim sRet As String
Dim ReturnValue As String

sRet = String(255, Chr(0))

ReturnValue = Left(sRet, getprivateprofilestring(section, ByVal KeyName$, "", sRet, Len(sRet), FileName))

If StrExt.IsNullOrWhiteSpace(ReturnValue) Then
    INIRead = DefaultValue
Else
    INIRead = ReturnValue
End If
End Function

Public Function INIWrite(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
Dim R
R = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
End Function
Public Function DownloadFile(url As String, LocalFilename As String) As Boolean
Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, url, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function
Public Function RandomNumber(Lowerbound As Integer, Upperbound As Integer) As Integer
RandomNumber = Int((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
End Function
Public Function GetPathFromFilename(File As String) As String
GetPathFromFilename = Left(File, Len(File) - Len(GetFileNameFromFilePath(File, True)))
End Function
Public Function GetCurrentUser() As String
Dim sUser As String
Dim lpBuff As String * 1024

GetUserName lpBuff, Len(lpBuff)
sUser = Left$(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
lpBuff = ""
    
GetCurrentUser = sUser
End Function

Public Function GetActiveWindowTitle() As String
Dim strTitle As String
Dim lngRet As Long

lngRet = GetForegroundWindow()
strTitle = String(GetWindowTextLength(lngRet) + 1, Chr$(0))
GetWindowText lngRet, strTitle, Len(strTitle)

GetActiveWindowTitle = Trim(strTitle)
End Function
Public Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal As Long, sShortPathName As String, iLen As Integer

sShortPathName = Space(255)
iLen = Len(sShortPathName)

lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
GetShortName = Left(sShortPathName, lRetVal)
End Function

Public Sub Pause(HowLong As Long)
Dim tick As Long
tick = GetTickCount()
Do
  DoEvents
Loop Until tick + HowLong < GetTickCount
End Sub
Public Function BrowseForFolder(ByVal lngHwnd As Long, ByVal strPrompt As String) As String
    On Error GoTo ehBrowseForFolder

    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo
    
    With udtBI
        .lngHwnd = lngHwnd
        .lpszTitle = lstrcat(strPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lngIDList = SHBrowseForFolder(udtBI)

    If lngIDList <> 0 Then
        
        strPath = String(MAX_PATH, 0)
        lngResult = SHGetPathFromIDList(lngIDList, strPath)
        
        Call CoTaskMemFree(lngIDList)
        
        intNull = InStr(strPath, vbNullChar)
        
        If intNull > 0 Then
            'Set the value
            strPath = Left(strPath, intNull - 1)
        End If
    End If
    
    BrowseForFolder = strPath
    Exit Function 'Abort

ehBrowseForFolder:
BrowseForFolder = Empty
End Function
Function LastFolderNameFromFullPath(ByVal sDir As String) As String
Dim splitDir
Dim c As Integer

splitDir = Split(sDir, "\")
c = UBound(splitDir)

LastFolderNameFromFullPath = splitDir(c)
End Function
Public Function DeleteDir(path As String)
Dim tempor As String

tempor = Dir(path + "\*.*", 22)

Do While tempor <> ""
    If tempor <> "." And tempor <> ".." Then
    If (GetAttr(path + "\" + tempor) And 16) = 0 Then
        SetAttr path + "\" + tempor, 0
        Kill path + "\" + tempor
    Else
        tempor = Dir(path + "\*.*", 22)
    End If
    End If
    tempor = Dir
Loop

RmDir path
End Function
Public Function GetFileNameFromFilePath(FilePath As String, WithExtension As Boolean) As String
Dim posn As Integer, I As Integer
Dim FName As String
    
posn = 0
For I = 1 To Len(FilePath)
    If (Strings.Mid(FilePath, I, 1) = "\") Then posn = I
Next I

FName = Strings.Right(FilePath, Len(FilePath) - posn)
posn = InStr(FName, ".")

If posn <> 0 Then
    If WithExtension = True Then
        FName = Strings.Left(FName, Len(FName))
    Else
        FName = Strings.Left(FName, posn - 1)
    End If
End If

GetFileNameFromFilePath = FName
End Function
