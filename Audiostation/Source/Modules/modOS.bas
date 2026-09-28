Attribute VB_Name = "modOS"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public VerMarjor As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Public Function GetWindowsVersion() As String
Dim osv As OSVERSIONINFO

osv.OSVSize = Len(osv)

If GetVersionEx(osv) = 1 Then
    Select Case osv.PlatformID
        Case VER_PLATFORM_WIN32s
            GetWindowsVersion = "Win32s on Windows 3.1"
        Case VER_PLATFORM_WIN32_NT
            GetWindowsVersion = "Windows NT"
            
            VerMarjor = osv.dwVerMajor
            
            Select Case osv.dwVerMajor
                Case 3
                    GetWindowsVersion = "Windows NT 3.5"
                Case 4
                    GetWindowsVersion = "Windows NT 4.0"
                Case 5
                    Select Case osv.dwVerMinor
                        Case 0
                            GetWindowsVersion = "Windows 2000"
                        Case 1
                            GetWindowsVersion = "Windows XP"
                        Case 2
                            GetWindowsVersion = "Windows Server 2003"
                    End Select
                Case 6
                    Select Case osv.dwVerMinor
                        Case 0
                            GetWindowsVersion = "Windows Vista/Server 2008"
                        Case 1
                            GetWindowsVersion = "Windows 7/Server 2008 R2"
                        Case 2
                            GetWindowsVersion = "Windows 8/Server 2012"
                        Case 3
                            GetWindowsVersion = "Windows 8.1/Server 2012 R2"
                    End Select
            End Select
    
        Case VER_PLATFORM_WIN32_WINDOWS:
            Select Case osv.dwVerMinor
                Case 0
                    GetWindowsVersion = "Windows 95"
                Case 90
                    GetWindowsVersion = "Windows Me"
                Case Else
                    GetWindowsVersion = "Windows 98"
            End Select
    End Select
Else
    GetWindowsVersion = "Unable to identify your version of Windows."
End If
End Function


