VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudstaService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Private Const WM_COPYDATA = &H4A
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long) 'Copies a block of memory from one location to another.
Private Sub Class_Initialize()
Shell App.path & "\audsta\audsta.exe", vbHide
End Sub
Public Function SendMessageToAudsta(message As String)
Dim cds As COPYDATASTRUCT
Dim ThWnd As Long
Dim buf(1 To 255) As Byte
Dim a As String
Dim I As Long

' Get the hWnd of the target application
ThWnd = FindWindow(vbNullString, "audsta")

' make sure window was found!
If ThWnd = 0 Then: MsgBox "Window Not found!": Exit Function

' string to send
a$ = message

' Copy the string into a byte array, converting it to ASCII
Call CopyMemory(buf(1), ByVal a$, Len(a$))

cds.dwData = 3 ' NOTE: USING 3 HERE AS AN IDENTIFYING MARKER!
cds.cbData = Len(a$) + 1
cds.lpData = VarPtr(buf(1))

I = SendMessage(ThWnd, WM_COPYDATA, Form_Main.hwnd, cds)
End Function
