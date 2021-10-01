VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SibraSoft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum enumTimeSerial
    [LongTimeSerial] = 0
    [SmallTimeSerial] = 1
End Enum

Public Enum convTo
    kb = 1
    MB = 2
    GB = 3
    TB = 4
End Enum

Public Enum enumMsgType
    [logError]
    [logWarning]
    [logInfo]
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

Enum SP
    [System Path]
    Desktop
    [Start Menu]
End Enum

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function SHChangeNotify Lib "shell32.dll" (ByVal wEventID As Long, ByVal uFlags As Long, ByVal dwItem1 As String, ByVal dwItems As String) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFilename As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function RtlGetVersion Lib "NTDLL" (ByRef lpVersionInformation As Long) As Long

'INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Function NativeGetVersion() As String
Dim tOSVw(&H54) As Long
tOSVw(0) = &H54 * &H4
Call RtlGetVersion(tOSVw(0))
NativeGetVersion = Join(Array(tOSVw(1), tOSVw(2), tOSVw(3)), ".")
End Function
Public Function GetCurrentWindowsVersion() As String
Dim VersionToName As String

Select Case NativeGetVersion
    Case "5.1": VersionToName = "Windows XP"
    Case "5.3": VersionToName = "Windows 2003 (SERVER)"
    Case "6.0": VersionToName = "Windows Vista"
    Case "6.1": VersionToName = "Windows 7"
    Case "6.2": VersionToName = "Windows 8"
    Case "6.3": VersionToName = "Windows 8.1"
    Case "10.0": VersionToName = "Windows 10"
    Case Else: VersionToName = "Unknown"
End Select

GetCurrentWindowsVersion = VersionToName
End Function
Public Sub TerminateProcessByPid(pid As Long)
Shell "taskkill.exe /f /t /pid " & pid, vbHide
End Sub
Public Sub TerminateProcessByName(Name As String)
Shell "taskkill.exe /f /t /im " & Name
End Sub
Public Sub RegUnReg(ByVal inFileSpec As String, Optional inHandle As String = "")
    On Error Resume Next
    Dim lLib As Long                 ' Store handle of the control library
    Dim lpDLLEntryPoint As Long      ' Store the address of function called
    Dim lpThreadID As Long           ' Pointer that receives the thread identifier
    Dim lpExitCode As Long           ' Exit code of GetExitCodeThread
    Dim mThread
    Dim mresult As Long
    
    lLib = LoadLibrary(inFileSpec)
    If lLib = 0 Then
        Debug.Print "Failure loading control DLL"
        Exit Sub
    End If
    
    If inHandle = "" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
    ElseIf inHandle = "U" Or inHandle = "u" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllUnregisterServer")
    Else
        Debug.Print "Unknown command handle"
        Exit Sub
    End If
    If lpDLLEntryPoint = vbNull Then
        GoTo earlyExit1
    End If
    
    Screen.MousePointer = vbHourglass
    
    mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
    If mThread = 0 Then
        GoTo earlyExit1
    End If
    
    mresult = WaitForSingleObject(mThread, 10000)
    If mresult <> 0 Then
        GoTo earlyExit2
    End If
    
    CloseHandle mThread
    FreeLibrary lLib
    
    Screen.MousePointer = vbDefault
    Debug.Print "Process completed"
    Exit Sub
    
    
earlyExit1:
    Screen.MousePointer = vbDefault
    Debug.Print "Process failed in obtaining entry point or creating thread."
    FreeLibrary lLib
    Exit Sub
    
earlyExit2:
    Screen.MousePointer = vbDefault
    Debug.Print "Process failed in signaled state or time-out."
    FreeLibrary lLib
     ' Terminate the thread to free up resources that are used by the thread
     ' NB Calling ExitThread for an application's primary thread will cause
     ' the application to terminate
    lpExitCode = GetExitCodeThread(mThread, lpExitCode)
    ExitThread lpExitCode
End Sub
Public Function ShellAndWait(Application As String, parameters As String) As Long
Dim AppPath As String
Dim oShell As Object

Set oShell = CreateObject("WSCript.shell")

AppPath = App.path & "\" & Application

If InStr(1, AppPath, " ") > 0 Then
    oShell.Run "cmd /C " & Chr(34) & App.path & "\" & Application & Chr(34) & " " & parameters, 0, True
Else
    oShell.Run "cmd /C " & App.path & "\" & Application & " " & parameters, 0, True
End If
End Function
Function IsOddNumber(ByVal iNum As Integer) As Boolean
    IsOddNumber = ((iNum \ 2) * 2 <> iNum)
End Function
Public Property Get ActiveWindow()
ActiveWindow = GetActiveWindow
End Property
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
Public Function StringBetween(FirstString As String, SecondString As String, TargetString As String) As String
Dim String1Pos As Integer
Dim String2Pos As Integer

String1Pos = InStr(1, TargetString, FirstString) + 1
String2Pos = InStr(1, Mid(TargetString, String1Pos), SecondString) - 1

StringBetween = Mid(TargetString, String1Pos, String2Pos)
End Function
Public Sub FilePutContents(file As String, Contents As String)
Dim iFileNo As Integer

iFileNo = FreeFile

Open file For Output As #iFileNo
    Print #iFileNo, Contents
Close #iFileNo
End Sub
Public Function FileGetContents(file As String) As String
Dim intFile As Integer
Dim strData As String

intFile = FreeFile

Open file For Input As #intFile
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
Public Function GetFileNameFromFile(flname As String) As String
Dim posn As Integer, I As Integer
Dim FName As String

posn = 0
For I = 1 To Len(flname)
    If (Mid(flname, I, 1) = "\") Then posn = I
Next I
FName = Right(flname, Len(flname) - posn)
posn = InStr(FName, ".")
    If posn <> 0 Then
        FName = Left(FName, posn - 1)
    End If
GetFileNameFromFile = FName
End Function
Public Function WriteEntryToWindowsEventLog(ErrorCode As Integer, MsgDescription As String, msgType As enumMsgType, msgLog As String)
Dim strErrorType As String

If msgType = logError Then: strErrorType = "ERROR"
If msgType = logWarning Then: strErrorType = "WARNING"
If msgType = logInfo Then: strErrorType = "INFORMATION"
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
Public Function TimeString(seconds As Long) As String

Dim lHrs As Long
Dim lMinutes As Long
Dim lSeconds As Long

lSeconds = seconds

lHrs = Int(lSeconds / 3600)
lMinutes = (Int(lSeconds / 60)) - (lHrs * 60)
lSeconds = Int(lSeconds Mod 60)

Dim sAns As String

If lSeconds = 60 Then
    lMinutes = lMinutes + 1
    lSeconds = 0
End If

If lMinutes = 60 Then
    lMinutes = 0
    lHrs = lHrs + 1
End If

sAns = format(CStr(lMinutes), "00") & ":" & format(CStr(lSeconds), "00")

TimeString = sAns
End Function
Public Function SecondsToTimeSerial(ByVal sec As Long, ReturnType As enumTimeSerial) As String
If ReturnType = LongTimeSerial Then
    SecondsToTimeSerial = format(DateAdd("s", sec, "00:00:00"), "HH:mm:ss")
Else
    SecondsToTimeSerial = format(DateAdd("s", sec, "00:00"), "mm:ss")
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
Public Function StringStartsWith(ByVal find_text As String, ByVal string_source As String, Optional ByVal caseSensitive As Boolean = True) As Boolean
If caseSensitive Then
    StringStartsWith = (Left$(string_source, LenB(find_text) / 2) = find_text)
Else
    StringStartsWith = (Left$(LCase(string_source), LenB(find_text) / 2) = LCase$(find_text))
End If
End Function
Public Function StringFormat(SourceString As String, ParamArray Values()) As String
Dim I As Integer
Dim ReturnValue As String

ReturnValue = SourceString

For I = 0 To UBound(Values)
    ReturnValue = Replace(ReturnValue, "{" & I & "}", Values(I))
Next

StringFormat = ReturnValue
End Function
Public Function Explode(TargetString As String, Delimiter As String, Posistion As Integer) As String
Dim SplitValue

SplitValue = Split(TargetString, Delimiter)

Explode = SplitValue(Posistion)
End Function
Public Function StringToCollection(TargetString As String, Seperator As String) As Collection
Dim I As Integer
Dim SplitValue

SplitValue = Split(TargetString, Seperator)

For I = 0 To UBound(SplitValue)
    StringToCollection.Add SplitValue(I)
Next
End Function
Public Function CollectionToString(TargetCollection As Collection, Seperator As String) As String
Dim I As Integer
Dim ReturnValue As String

For I = 1 To TargetCollection.Count
    ReturnValue = ReturnValue & TargetCollection(I) & Seperator
Next

CollectionToString = ReturnValue
End Function
Public Function INIRead(Section As String, KeyName As String, FileName As String, Optional DefaultValue As String) As String
Dim sRet As String
Dim ReturnValue As String

sRet = String(255, Chr(0))

ReturnValue = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))

If ReturnValue = vbNullString Then
    INIRead = DefaultValue
Else
    INIRead = ReturnValue
End If
End Function

Public Function INIWrite(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
Dim r
r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function
Public Function DownloadFile(url As String, LocalFilename As String) As Boolean
Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, url, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function

Public Function cBytes(lBytes As Long, convertto As convTo) As Double
cBytes = lBytes / (1024 ^ convertto)
End Function
Public Function RandomNumber(Lowerbound As Integer, Upperbound As Integer) As Integer
RandomNumber = Int((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
End Function
Public Function GetLongFilename(ByVal sShortFilename As String) As String
Dim lRet As Long
Dim sLongFileName As String
   sLongFileName = String$(1024, " ")
   lRet = GetLongPathName(sShortFilename, sLongFileName, Len(sLongFileName))


If lRet > Len(sLongFileName) Then
   sLongFileName = String$(lRet + 1, " ")
   lRet = GetLongPathName(sShortFilename, sLongFileName, Len(sLongFileName))
End If


If lRet > 0 Then
   GetLongFilename = Left$(sLongFileName, lRet)
End If
End Function
Public Function cUNC(sPathName As String) As String
Dim szValue As String, szValueName As String, sUNCName As String
Dim lErrCode As Long, lEndBuffer As Long

Const lLenUNC As Long = 520
Const NO_ERROR As Long = 0
Const ERROR_NOT_CONNECTED As Long = 2250
Const ERROR_BAD_DEVICE = 1200&
Const ERROR_MORE_DATA = 234
Const ERROR_CONNECTION_UNAVAIL = 1201&
Const ERROR_NO_NETWORK = 1222&
Const ERROR_EXTENDED_ERROR = 1208&
Const ERROR_NO_NET_OR_BAD_PATH = 1203&

'Verify whether the disk is connected to the network
If Mid$(sPathName, 2, 1) = ":" Then
    sUNCName = String$(lLenUNC, 0)
    lErrCode = WNetGetConnection(Left$(sPathName, 2), sUNCName, lLenUNC)
    lEndBuffer = InStr(sUNCName, vbNullChar) - 1
    'Can ignore the errors below (will still return the correct UNC)
    If lEndBuffer > 0 And (lErrCode = NO_ERROR Or lErrCode = ERROR_CONNECTION_UNAVAIL Or lErrCode = ERROR_NOT_CONNECTED) Then
        'Success
        sUNCName = Trim$(Left$(sUNCName, InStr(sUNCName, vbNullChar) - 1))
        cUNC = sUNCName & Mid$(sPathName, 3)
    Else
        'Error, return original path
        cUNC = sPathName
    End If
Else
    'Already a UNC Path
    cUNC = sPathName
End If
End Function
Public Function GetCurrentUser() As String
Dim sUser As String
Dim lpBuff As String * 1024

GetUserName lpBuff, Len(lpBuff)
sUser = Left$(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
lpBuff = ""
    
GetCurrentUser = sUser
End Function
Public Function CreateShortCut(ByVal TargetPath As String, ByVal ShortCutPath As SP, ByVal ShortCutname As String, Optional ByVal WorkPath As String, Optional ByVal Window_Style As Integer, Optional ByVal IconNum As Integer)
Dim VbsObj As Object
Set VbsObj = CreateObject("WScript.Shell")
Dim MyShortcut As Object

If ShortCutPath = [System Path] Then: ShortCutPath = "C:\windows\"

ShortCutPath = VbsObj.SpecialFolders(ShortCutPath)
Set MyShortcut = VbsObj.CreateShortCut(ShortCutPath & ShortCutname & ".lnk")

MyShortcut.TargetPath = TargetPath
MyShortcut.WorkingDirectory = WorkPath
MyShortcut.WindowStyle = Window_Style
MyShortcut.IconLocation = TargetPath & "," & IconNum
MyShortcut.Save
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
Dim Tick As Long
Tick = GetTickCount()
Do
  DoEvents
Loop Until Tick + HowLong < GetTickCount
End Sub
Public Function cChar(ByVal Text As String, ByVal char As String) As Long
Dim v As Variant
v = Split(Text, char)

cChar = UBound(v)
End Function
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
Dim C As Integer

splitDir = Split(sDir, "\")
C = UBound(splitDir)

LastFolderNameFromFullPath = splitDir(C)
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
