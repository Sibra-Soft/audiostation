Attribute VB_Name = "modExternalApp"
Option Explicit

Public Const WM_VSCROLL As Long = &H115
Public Const WM_HSCROLL As Long = &H114
Public Const WM_CLOSE As Long = &H10

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type TChildWindowInfo
    hwnd As Long
    ClassName As String
    Text As String
    style As Long
    HasVScroll As Boolean
    HasHScroll As Boolean
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Const BM_CLICK As Long = &HF5
Public Const NORMAL_PRIORITY_CLASS As Long = &H20&

Public g_SearchCaption As String
Public g_FoundChild As Long
Public g_TargetPID As Long
Public g_FoundWindow As Long
Public g_RequireVisible As Boolean
Public g_FoundScrollBar As Long
Public g_ChildInfos() As TChildWindowInfo
Public g_ChildCount As Long

Public g_pi As PROCESS_INFORMATION

Public Const STARTF_USESHOWWINDOW As Long = &H1
Public Const SW_HIDE As Long = 0
Public Const SW_SHOWNORMAL As Long = 1

Public Const GWL_STYLE As Long = -16

Public Const WS_VSCROLL As Long = &H200000
Public Const WS_HSCROLL As Long = &H100000

Public Const SB_LINELEFT As Long = 0
Public Const SB_LINERIGHT As Long = 1
Public Const SB_LINEDOWN As Long = 1
Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim cls As String
    Dim txt As String
    Dim n As Long

    cls = String$(256, vbNullChar)
    n = GetClassName(hwnd, cls, 255)
    If n > 0 Then cls = Left$(cls, n)

    txt = String$(512, vbNullChar)
    n = GetWindowText(hwnd, txt, 511)
    If n > 0 Then
        txt = Left$(txt, n)
    Else
        txt = ""
    End If

    If StrComp(cls, "Button", vbTextCompare) = 0 Then
        If StrComp(txt, g_SearchCaption, vbTextCompare) = 0 Then
            g_FoundChild = hwnd
            EnumChildProc = 0
            Exit Function
        End If
    End If

    EnumChildProc = 1
End Function

Public Function FindButtonRecursive(ByVal hWndParent As Long, ByVal ButtonCaption As String) As Long
    g_SearchCaption = ButtonCaption
    g_FoundChild = 0

    EnumChildWindows hWndParent, AddressOf EnumChildProc, 0

    FindButtonRecursive = g_FoundChild
End Function

Public Function ClickButtonInWindow(ByVal hWndMain As Long, ByVal ButtonCaption As String) As Boolean
    Dim hBtn As Long

    If hWndMain = 0 Then Exit Function

    hBtn = FindButtonRecursive(hWndMain, ButtonCaption)
    If hBtn = 0 Then Exit Function

    SendMessage hBtn, BM_CLICK, 0, 0
    ClickButtonInWindow = True
End Function
Public Function EnumFindScrollBarProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim cls As String

    cls = GetWindowClass(hwnd)

    If StrComp(cls, "ScrollBar", vbTextCompare) = 0 Then
        g_FoundScrollBar = hwnd
        EnumFindScrollBarProc = 0
        Exit Function
    End If

    EnumFindScrollBarProc = 1
End Function

Public Function FindFirstScrollBar(ByVal hWndParent As Long) As Long
    g_FoundScrollBar = 0
    EnumChildWindows hWndParent, AddressOf EnumFindScrollBarProc, 0
    FindFirstScrollBar = g_FoundScrollBar
End Function
Public Function GetWindowClass(ByVal hwnd As Long) As String
    Dim buf As String
    Dim n As Long

    buf = String$(256, vbNullChar)
    n = GetClassName(hwnd, buf, 255)

    If n > 0 Then
        GetWindowClass = Left$(buf, n)
    Else
        GetWindowClass = ""
    End If
End Function

Public Sub AddChildInfo(ByVal hwnd As Long, ByVal cls As String, ByVal txt As String, ByVal style As Long)
    g_ChildCount = g_ChildCount + 1
    ReDim Preserve g_ChildInfos(1 To g_ChildCount)

    g_ChildInfos(g_ChildCount).hwnd = hwnd
    g_ChildInfos(g_ChildCount).ClassName = cls
    g_ChildInfos(g_ChildCount).Text = txt
    g_ChildInfos(g_ChildCount).style = style
    g_ChildInfos(g_ChildCount).HasVScroll = ((style And WS_VSCROLL) <> 0)
    g_ChildInfos(g_ChildCount).HasHScroll = ((style And WS_HSCROLL) <> 0)
End Sub
Public Function EnumChildCollectProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim cls As String
    Dim txt As String
    Dim style As Long

    cls = GetWindowClass(hwnd)
    txt = GetControlText(hwnd)
    style = GetWindowLong(hwnd, GWL_STYLE)

    AddChildInfo hwnd, cls, txt, style

    EnumChildCollectProc = 1
End Function
Public Sub CollectChildWindows(ByVal hWndParent As Long)
    g_ChildCount = 0
    Erase g_ChildInfos

    EnumChildWindows hWndParent, AddressOf EnumChildCollectProc, 0
End Sub
Public Function FindFirstStatic(ByVal hWndParent As Long) As Long
FindFirstStatic = FindWindowEx(hWndParent, 0, "Static", vbNullString)
End Function
Public Function FindNextStatic(ByVal hWndParent As Long, ByVal hWndAfter As Long) As Long
FindNextStatic = FindWindowEx(hWndParent, hWndAfter, "Static", vbNullString)
End Function
Public Function EnumStaticProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim cls As String
    Dim txt As String
    Dim n As Long

    cls = String$(256, vbNullChar)
    n = GetClassName(hwnd, cls, 255)
    If n > 0 Then cls = Left$(cls, n)

    If StrComp(cls, "Static", vbTextCompare) = 0 Then
        txt = GetControlText(hwnd)
        Debug.Print "Static hWnd=" & hwnd & " | Text=" & txt
    End If

    EnumStaticProc = 1
End Function

Public Function GetControlText(ByVal hwnd As Long) As String
    Dim buf As String
    Dim n As Long

    If hwnd = 0 Then Exit Function

    buf = String$(1024, vbNullChar)
    n = GetWindowText(hwnd, buf, Len(buf) - 1)

    If n > 0 Then
        GetControlText = Left$(buf, n)
    Else
        GetControlText = ""
    End If
End Function
Public Function ScrollBarStepBack(ByVal hScrollBar As Long) As Boolean
    Dim hParent As Long

    If hScrollBar = 0 Then Exit Function

    hParent = GetParent(hScrollBar)
    If hParent = 0 Then Exit Function

    SendMessage hParent, WM_HSCROLL, SB_LINELEFT, hScrollBar
    ScrollBarStepBack = True
End Function
Public Function ScrollBarStepForward(ByVal hScrollBar As Long) As Boolean
    Dim hParent As Long

    If hScrollBar = 0 Then Exit Function

    hParent = GetParent(hScrollBar)
    If hParent = 0 Then Exit Function

    SendMessage hParent, WM_HSCROLL, SB_LINERIGHT, hScrollBar
    ScrollBarStepForward = True
End Function
Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim pid As Long

    GetWindowThreadProcessId hwnd, pid

    If pid = g_TargetPID Then
        If g_RequireVisible Then
            If IsWindowVisible(hwnd) = 0 Then
                EnumWindowsProc = 1
                Exit Function
            End If
        End If

        g_FoundWindow = hwnd
        EnumWindowsProc = 0
        Exit Function
    End If

    EnumWindowsProc = 1
End Function
Public Function StartExternalApp(ExePath As String, Optional Args As String = "", Optional StartHidden As Boolean = False) As Boolean
Dim si As STARTUPINFO
Dim cmd As String
Dim rc As Long

si.cb = Len(si)

If StartHidden Then
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
Else
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_SHOWNORMAL
End If

If Len(Trim$(Args)) > 0 Then
    cmd = """" & ExePath & """ " & Args
Else
    cmd = """" & ExePath & """"
End If

Debug.Print ExePath
rc = CreateProcess(ExePath, cmd, 0, 0, 0, NORMAL_PRIORITY_CLASS, 0, App.path, si, g_pi)

StartExternalApp = (rc <> 0)
End Function
Public Function CloseExternalApp(ByVal hWndMain As Long) As Boolean
If hWndMain = 0 Then Exit Function

SendMessage hWndMain, WM_CLOSE, 0, 0
CloseExternalApp = True
End Function
Public Sub CleanupExternalApp()
If g_pi.hThread <> 0 Then
    CloseHandle g_pi.hThread
    g_pi.hThread = 0
End If

If g_pi.hProcess <> 0 Then
    CloseHandle g_pi.hProcess
    g_pi.hProcess = 0
End If

g_pi.dwProcessId = 0
g_pi.dwThreadId = 0
End Sub
Public Function WaitForMainWindowByPID(pid As Long, Optional TimeoutMs As Long = 10000, Optional RequireVisible As Boolean = True) As Long
Dim t As Long

g_TargetPID = pid
g_FoundWindow = 0
g_RequireVisible = RequireVisible

Do While t < TimeoutMs
    EnumWindows AddressOf EnumWindowsProc, 0

    If g_FoundWindow <> 0 Then
        WaitForMainWindowByPID = g_FoundWindow
        Exit Function
    End If

    Sleep 100
    t = t + 100
Loop
End Function
