Attribute VB_Name = "Windows"
Enum SP
    [System Path]
    Desktop
    [Start Menu]
End Enum

Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare Function GetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function SHChangeNotify Lib "shell32.dll" (ByVal wEventID As Long, ByVal uFlags As Long, ByVal dwItem1 As String, ByVal dwItems As String) As Long
Private Declare Function GetLongPathName Lib "Kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0
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
Public Function ConvertToUNC(sPathName As String) As String
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
        ConvertToUNC = sUNCName & Mid$(sPathName, 3)
    Else
        'Error, return original path
        ConvertToUNC = sPathName
    End If
Else
    'Already a UNC Path
    ConvertToUNC = sPathName
End If
End Function
Public Sub AssociateFile(ByVal sAppName As String, ByVal sEXE As String, ByVal sExt As String, Optional ByVal sCommand As String, Optional ByVal sIcon As String)
Dim sCommandString As String
Dim lRegKey As Long
        
Call RegCreateKey(HKEY_CLASSES_ROOT, "." & sExt, lRegKey)
Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sAppName, Len(sAppName))
Call RegCloseKey(lRegKey)
       
sCommand = "\Shell\" & IIf(Len(sCommand), sCommand, "Open") & "\Command"
Call RegCreateKey(HKEY_CLASSES_ROOT, sAppName & sCommand, lRegKey)
Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sEXE, Len(sEXE))
Call RegCloseKey(lRegKey)
    
If Len(sIcon) Then
    Call RegCreateKey(HKEY_CLASSES_ROOT, sAppName & "\DefaultIcon", lRegKey)
    Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sIcon, Len(sIcon))
    Call RegCloseKey(lRegKey)
End If

SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, vbNullString, vbNullString
End Sub
Public Function GetCurrentUser() As String
Dim sUser As String
Dim lpBuff As String * 1024

GetUserName lpBuff, Len(lpBuff)
sUser = Left$(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
lpBuff = ""
    
GetCurrentUser = sUser
End Function
Public Function GetWindowsVersion() As String
strComputer = "."

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOperatingSystem In colOperatingSystems
    Get_Windows_Version = objOperatingSystem.Caption
Next
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
