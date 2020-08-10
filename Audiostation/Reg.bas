Attribute VB_Name = "Reg"
Option Explicit

Private Const ERROR_NONE = 0

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
"RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
"RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As String, lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, lpData As _
Long, lpcbData As Long) As Long

Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As Long, lpcbData As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
String, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
ByVal cbData As Long) As Long

Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" _
    (ByVal hKey As Long, ByVal pszSubKey As String) As Long
Private Function SetValueEx(ByVal hKey As Long, sValueName As String, _
lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
                                           lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, _
lType, lValue, 4)
        End Select
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
      String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        Case REG_SZ:  ' For strings
            sValue = String(cch, 0)

            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
                     sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        Case REG_DWORD:  ' For DWORDS
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
                  lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function

QueryValueExError:
    Resume QueryValueExExit
End Function

'Example Call: CreateNewKey HKEY_LOCAL_MACHINE, "TestKey"
'Example Call: CreateNewKey HKEY_LOCAL_MACHINE, "TestKey\SubKey1\SubKey2"
Public Sub CreateNewKey(hType As Long, sNewKeyName As String, lRetVal As Long)
 Dim hNewKey As Long         'handle to the new key
'    Dim lRetVal As Long         'result of the RegCreateKeyEx function

 lRetVal = RegCreateKeyEx(hType, sNewKeyName, 0&, _
         vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
         0&, hNewKey, lRetVal)
 RegCloseKey (hNewKey)
End Sub

'Example Call:  SetKeyValue HKEY_CURRENT_USER, "TestKey\SubKey1", "StringValue", "Hello", REG_SZ(REG_BINARY)
Public Sub SetKeyValue(hType As Long, sKeyName As String, sValueName As String, _
vValueSetting As Variant, lValueType As Long, lRetVal As Long)
'    Dim lRetVal As Long      'result of the SetValueEx function
 Dim hKey As Long         'handle of open key

 'open the specified key
 lRetVal = RegOpenKeyEx(hType, sKeyName, 0, KEY_SET_VALUE, hKey)
 lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
 RegCloseKey (hKey)
End Sub

'Example Call: QueryValue "TestKey\SubKey1", "StringValue"
Public Sub QueryValue(hType As Long, sKeyName As String, sValueName As String, vValue As Variant, lRetVal As Long)
'       Dim lRetVal As Long      'result of the API functions
    Dim hKey As Long         'handle of opened key
'       Dim vValue As Variant      'setting of queried value

    lRetVal = RegOpenKeyEx(hType, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
'MsgBox vValue
    RegCloseKey (hKey)
End Sub

'Using Recursive System Version So Deleting Parent Deletes Children.
'Example Call: deleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "test"
Public Sub DeleteKeys(hType As Long, subKey As String, Key As String, lResult As Long)
 'Dim lResult As Long
 Dim hKey As Long
 If RegOpenKeyEx(hType, subKey, 0, KEY_ALL_ACCESS, hKey) = ERROR_NONE Then
     lResult = SHDeleteKey(hKey, Key)
     RegCloseKey hKey
 End If
End Sub

