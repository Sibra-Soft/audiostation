VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public DeviceCount As Integer
Public Devices As New Collection
Public Function FindDevice(DeviceName As String) As Boolean
Dim I As Integer

For I = 1 To Devices.Count
    If Devices.Item(I) = DeviceName Then
        FindDevice = True
        Exit Function
    End If
Next
End Function
Public Sub GetDeviceList()
Dim DevCount As Integer

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SoundDevice", , 48)

For Each objItem In colItems
    Devices.Add objItem.Caption
    
    DevCount = DevCount + 1
Next

DeviceCount = DevCount
End Sub
