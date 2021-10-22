VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistrySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const REG_SZ As Long = 1

Private Const HKEY_CURRENT_USER = &H80000001
Private Const BASE_KEY As String = "SOFTWARE"

Private Const ERROR_NONE As Long = 0
Private Const ERROR_KEY_DOES_NOT_EXIST As Long = 2

Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_CREATE_LINK As Long = &H20
Private Const SYNCHRONIZE As Long = &H100000

Private Const KEY_ALL_ACCESS  As Long = ((STANDARD_RIGHTS_ALL Or _
                                            KEY_QUERY_VALUE Or _
                                            KEY_SET_VALUE Or _
                                            KEY_CREATE_SUB_KEY Or _
                                            KEY_ENUMERATE_SUB_KEYS Or _
                                            KEY_NOTIFY Or _
                                            KEY_CREATE_LINK) _
                                            And (Not SYNCHRONIZE))
Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or _
                                            KEY_QUERY_VALUE Or _
                                            KEY_ENUMERATE_SUB_KEYS Or _
                                            KEY_NOTIFY) _
                                            And (Not SYNCHRONIZE))

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Sub WriteSetting(ByVal sCompanyName As String, ByVal sAppName As String, ByVal sKey As String, ByVal sSetting As String)
                    
Dim lRetVal         As Long
Dim sNewKey         As String
Dim lDisposition    As Long
Dim lHandle         As Long
Dim lErrNumber      As Long
Dim sErrDescription As String
Dim sErrSource      As String

On Error GoTo ERROR_HANDLER

If Trim(sCompanyName) = "" Then
    Err.Raise vbObjectError + 1000, , "AppName may not be empty"
End If
If Trim(sAppName) = "" Then
    Err.Raise vbObjectError + 1001, , "Section may not be empty"
End If
If Trim(sKey) = "" Then
    Err.Raise vbObjectError + 1002, , "Key may not be empty"
End If

sNewKey = BASE_KEY & "\" & Trim(sCompanyName) & "\" & Trim(sAppName)

' Create the key or open it if it already exists
lRetVal = RegCreateKeyEx(HKEY_CURRENT_USER, sNewKey, 0, vbNullString, 0, _
    KEY_ALL_ACCESS, 0, lHandle, lDisposition)
    
If lRetVal <> ERROR_NONE Then
    Err.Raise vbObjectError + 2000 + lRetVal, , _
        "Could not open/create registry section"
End If

' Set the key value
lRetVal = RegSetValueExString(lHandle, sKey, 0, REG_SZ, sSetting, _
    Len(sSetting))

If lRetVal <> ERROR_NONE Then
    Err.Raise vbObjectError + 2000 + lRetVal, , "Could not set key value"
End If

TIDY_UP:
    On Error Resume Next
    
    RegCloseKey lHandle
    
    If lErrNumber <> 0 Then
        On Error GoTo 0
        
        Err.Raise lErrNumber, sErrSource, sErrDescription
    End If
Exit Sub

ERROR_HANDLER:
    lErrNumber = Err.Number
sErrSource = Err.Source
sErrDescription = Err.Description
Resume TIDY_UP
End Sub
Public Function ReadSetting(ByVal sCompanyName As String, ByVal sAppName As String, ByVal sKey As String, Optional ByVal sDefault As String) As String
Dim lRetVal         As Long
Dim sFullKey        As String
Dim lHandle         As Long
Dim lType           As Long
Dim lLength         As Long
Dim sValue          As String
Dim lErrNumber      As Long
Dim sErrDescription As String
Dim sErrSource      As String

On Error GoTo ERROR_HANDLER

If Trim(sCompanyName) = "" Then
    Err.Raise vbObjectError + 1000, , "AppName may not be empty"
End If
If Trim(sAppName) = "" Then
    Err.Raise vbObjectError + 1001, , "Section may not be empty"
End If
If Trim(sKey) = "" Then
    Err.Raise vbObjectError + 1002, , "Key may not be empty"
End If

sFullKey = BASE_KEY & "\" & Trim(sCompanyName) & "\" & Trim(sAppName)

' Open up the key
lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sFullKey, 0, KEY_READ, lHandle)
If lRetVal <> ERROR_NONE Then
    If lRetVal = ERROR_KEY_DOES_NOT_EXIST Then
        ReadSetting = sDefault
        Exit Function
    Else
        Err.Raise vbObjectError + 2000 + lRetVal, , _
            "Could not open registry section"
    End If
End If

' Get size and type
lRetVal = RegQueryValueExNULL(lHandle, sKey, 0, lType, 0, lLength)
If lRetVal <> ERROR_NONE Then
    ReadSetting = sDefault
    Exit Function
End If

' Is it stored as a string in the registry?
If lType = REG_SZ Then
    sValue = String(lLength, 0)
    
    If lLength = 0 Then
        ReadSetting = ""
    Else
        lRetVal = RegQueryValueExString(lHandle, sKey, 0, lType, _
            sValue, lLength)
        
        If lRetVal = ERROR_NONE Then
            ReadSetting = Left(sValue, lLength - 1)
        Else
            ReadSetting = sDefault
        End If
    End If
Else
    Err.Raise vbObjectError + 2000 + lType, , _
        "Registry data not a string"
End If
    
TIDY_UP:
    On Error Resume Next
    
    RegCloseKey lHandle
    
    If lErrNumber <> 0 Then
        On Error GoTo 0
        
        Err.Raise lErrNumber, sErrSource, sErrDescription
    End If
Exit Function

ERROR_HANDLER:
    lErrNumber = Err.Number
    sErrSource = Err.Source
    sErrDescription = Err.Description
    Resume TIDY_UP
End Function

