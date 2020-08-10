Attribute VB_Name = "ModFileCommands"
Option Explicit

' Windows Registry Root Key Constants.
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

' Windows Registry Key Type Constants.
Public Const REG_OPTION_NON_VOLATILE = 0        ' Key is preserved when system is rebooted
Public Const REG_DWORD = 4                      ' 32-bit number

Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_SZ = 1                         ' Unicode nul terminated string

Public Const REG_BINARY = 3                     ' Free form binary

Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)

' Function Error Constants.
Public Const ERROR_SUCCESS = 0
Public Const ERROR_REG = 1

' Registry Access Rights.
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

' Windows Registry API Declarations.
' Registry API To Open A Key.
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long

' Registry API To Create A New Key.
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
  ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
  ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

' Registry API To Query A String Value.
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
  lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
  ' Note that if you declare the lpData parameter as String, you must pass it By Value.

' Registry API To Query A Long (DWORD) Value.
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
  lpType As Long, lpData As Long, lpcbData As Long) As Long

' Registry API To Query A NULL Value.
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
  lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

' Registry API To Set A String Value.
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
  ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
  ' Note that if you declare the lpData parameter as String, you must pass it By Value.

' Registry API To Set A Long (DWORD) Value.
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
  ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

' Registry API To Delete A Key.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long

' Registry API To Delete A Key Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
  (ByVal hKey As Long, ByVal lpValueName As String) As Long

' Registry API To Close A Key.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

' Constants For Error Messages.
Public Const OpenErr = "Error: Opening Registry Key!"
Public Const DeleteErr = "Error: Deleteing Key!"
Public Const CreateErr = "Error: Creating Key!"
Public Const QueryErr = "Error: Querying Value!"
Public Function FileCommand(Extension As String, Action As String, Command As String)


  Dim lRtn    As Long     ' API Return Code
  Dim hKey    As Long     ' Handle Of Open Key
  Dim lCdata  As Long     ' The Data
  Dim lValue  As Long     ' Long (DWORD) Value
  Dim sValue  As String   ' String Value
  Dim lRtype  As Long     ' Type Returned String Or DWORD
  Dim KeyName As String
  Dim lsize As Long
  

 ' Open The Registry Key.
  lRtn = RegOpenKeyEx(HKEY_CLASSES_ROOT, Extension, 0&, KEY_ALL_ACCESS, hKey)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox OpenErr
    RegCloseKey (hKey)
    Exit Function
  End If
  
  ' Query Registry Key For Value Type.
  lRtn = RegQueryValueExNULL(hKey, "", 0&, lRtype, 0&, lCdata)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox QueryErr
    RegCloseKey (hKey)
    Exit Function
  End If
  

      sValue = String(lCdata, 0)
      ' Get Registry String Value.
      lRtn = RegQueryValueExString(hKey, "", 0&, lRtype, sValue, lCdata)
  

  
  ' Close The Registry Key.
  RegCloseKey (hKey)


  
  'MsgBox (sValue)
  sValue = Left$(sValue, (Len(sValue) - 1))
  
  'RegCloseKey (hKey)
  KeyName = sValue + "\shell\" + Action
  'MsgBox (KeyName)
    
  ' Create The New Registry Key.
  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox CreateErr
  End If
  


  
  sValue = Action        ' Assign Key Value
  lsize = Len(sValue)      ' Get Size Of String
  ' Set String Value.
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, sValue, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
     MsgBox "Error Setting String Value!"
     RegCloseKey (hKey)
     Exit Function
  End If


  ' Close The Registry Key.
  RegCloseKey (hKey)
  
   KeyName = KeyName + "\command"
     
    
  ' Create The New Registry Key.
  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox CreateErr
  End If
  
    sValue = Command        ' Assign Key Value
  lsize = Len(sValue)      ' Get Size Of String
  ' Set String Value.
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, sValue, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
     MsgBox "Error Setting String Value!"
     RegCloseKey (hKey)
     Exit Function
  End If


  ' Close The Registry Key.
  RegCloseKey (hKey)
  
  
End Function
Public Function AssociateFile(Extension As String, Application As String, Identifier As String, Description As String, Icon As String)

  Dim lRtn    As Long     ' Returned Value From API Registry Call
  Dim hKey    As Long     ' Handle Of Open Key
  Dim lValue  As Long     ' Setting A Long Data Value
  Dim sValue  As String   ' Setting A String Data Value
  Dim lsize   As Long     ' Size Of String Data To Set
  Dim commandline As String
  

  ' Create The New Registry Key, the file extension
  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Extension, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox CreateErr
  End If
      
  lsize = Len(Identifier)      ' Get Size Of identifier String
  ' Set "(Default)" String Value to identifier
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, Identifier, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
      MsgBox "Error Setting String Value!"
      RegCloseKey (hKey)
      Exit Function
  End If

  ' Close The Registry Key.
  RegCloseKey (hKey)

  ' Create The New Registry Key, the file extension identifier
  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Identifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox CreateErr
  End If
  
    lsize = Len(Description)      ' Get Size Of file type description String
  ' Set (Default) String Value to description of the file type
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, Description, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
      MsgBox "Error Setting String Value!"
      RegCloseKey (hKey)
      Exit Function
  End If

  ' Close The Registry Key.
  RegCloseKey (hKey)


  ' Create The New Registry Key, the default icon key within the identifier key
  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, (Identifier + "\DefaultIcon"), 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox CreateErr
  End If
  
    lsize = Len(Icon)      ' Get Size Of String
  ' Set (Default) String Value to the full path name of the icon that will be associated with
  '    this file type
  
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, Icon, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
      MsgBox "Error Setting String Value!"
      RegCloseKey (hKey)
      Exit Function
  End If

  ' Close The Registry Key.
  RegCloseKey (hKey)



Identifier = Identifier + "\shell"
  ' Create The New Registry Key, the "shell" key within the identifier key

  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Identifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox CreateErr
  End If
  
  ' Close The Registry Key.
  RegCloseKey (hKey)


Identifier = Identifier + "\open"
  ' Create The New Registry Key, the "open" command key within the shell key

  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Identifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox CreateErr
  End If
  
  ' Close The Registry Key.
  RegCloseKey (hKey)


Identifier = Identifier + "\command"
  ' Create The New Registry Key, the "command"  key within the "open" command key

  lRtn = RegCreateKeyEx(HKEY_CLASSES_ROOT, Identifier, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                          KEY_ALL_ACCESS, 0&, hKey, lRtn)
  
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
    MsgBox CreateErr
  End If

    commandline = (Chr$(34) + Application + Chr$(34) + " " + Chr$(34) + "%1" + Chr$(34))
    lsize = Len(commandline)      ' Get Size Of String
  ' Set (Default) String Value of the "command" key to the command line to be used to open the file
  lRtn = RegSetValueExString(hKey, "", 0&, REG_SZ, commandline, lsize)
  ' Check For An Error.
  If lRtn <> ERROR_SUCCESS Then
      MsgBox "Error Setting String Value!"
      RegCloseKey (hKey)
      Exit Function
  End If

  ' Close The Registry Key.
  RegCloseKey (hKey)
End Function

