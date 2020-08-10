Attribute VB_Name = "RegisterLibrary"
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
    (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long

Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" _
  (ByVal lpLibFileName As String) As Long
  
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long

Private Declare Function CreateThread Lib "Kernel32" (lpThreadAttributes As Any, _
   ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, _
   ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
   
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
   
Private Declare Function GetExitCodeThread Lib "Kernel32" (ByVal hThread As Long, _
    lpExitCode As Long) As Long

Private Declare Sub ExitThread Lib "Kernel32" (ByVal dwExitCode As Long)

Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Sub RegUnReg(ByVal inFileSpec As String, Optional inHandle As String = "")
    On Error Resume Next
    Dim lLib As Long                 ' Store handle of the control library
    Dim lpDLLEntryPoint As Long      ' Store the address of function called
    Dim lpThreadID As Long           ' Pointer that receives the thread identifier
    Dim lpExitCode As Long           ' Exit code of GetExitCodeThread
    Dim mThread
    
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
