VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdioMasterVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLSCTX_INPROC_SERVER As Long = &H1&
Private Const CLSCTX_INPROC_HANDLER As Long = &H2&
Private Const CLSCTX_LOCAL_SERVER As Long = &H4&
Private Const CLSCTX_REMOTE_SERVER As Long = &H10&
Private Const CLSCTX_ALL As Long = CLSCTX_INPROC_SERVER Or _
                                   CLSCTX_INPROC_HANDLER Or _
                                   CLSCTX_LOCAL_SERVER Or _
                                   CLSCTX_REMOTE_SERVER


'
Private Const CLSID_MyContext As String = "{5929DCA2-52B8-449B-8A58-FDF6750634DA}"
Private Const CLSID_MMDeviceEnumerator As String = "{BCDE0395-E52F-467C-8E3D-C4579291692E}"
Private Const IID_IMMDeviceEnumerator As String = "{A95664D2-9614-4F35-A746-DE8DB63617E6}"
Private Const IID_IAudioEndpointVolume As String = "{5CDF2C82-841E-4546-9722-0CF74078229A}"

'
Private Type UUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type
 
'
Private Enum ERole
eConsole
eMultimedia
eCommunications
ERole_enum_count
End Enum

Private Enum EDataFlow
eRender
eCapture
eAll
EDataFlow_enum_count
End Enum


Private Type IMMDeviceEnumeratorVtbl '/* C style interface */
' IUnknown Interface
QueryInterface As Long
AddRef As Long
Release As Long
' IMMDeviceEnumerator Interface
EnumAudioEndpoints As Long
GetDefaultAudioEndpoint As Long
GetDevice As Long
RegisterEndpointNotificationCallback As Long
UnregisterEndpointNotificationCallback As Long
End Type


Private Type IMMDeviceVtbl '/* C style interface */
' IUnknown Interface
QueryInterface As Long
AddRef As Long
Release As Long
' IMMDevice Interface
Activate As Long
OpenPropertyStore As Long
GetId As Long
GetState As Long
End Type


Private Type IAudioEndpointVolumeVtbl '/* C style interface */
' IUnknown Interface
QueryInterface As Long
AddRef As Long
Release As Long
' IAudioEndpointVolume Interface
RegisterControlChangeNotify As Long
UnregisterControlChangeNotify As Long
GetChannelCount As Long
SetMasterVolumeLevel As Long
SetMasterVolumeLevelScalar As Long
GetMasterVolumeLevel As Long
GetMasterVolumeLevelScalar As Long
SetChannelVolumeLevel As Long
SetChannelVolumeLevelScalar As Long
GetChannelVolumeLevel As Long
GetChannelVolumeLevelScalar As Long
SetMute As Long
GetMute As Long
GetVolumeStepInfo As Long
VolumeStepUp As Long
VolumeStepDown As Long
QueryHardwareSupport As Long
GetVolumeRange As Long
End Type

'
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByRef dest As Any, ByRef Source As Any, ByVal bytes As Long)
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpszIID As Long, ByRef iid As UUID) As Long
Private Declare Function CallWindowProcA Lib "user32.dll" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function CoCreateInstance Lib "ole32.dll" (ByRef rclsid As UUID, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As UUID, ByRef ppv As Long) As Long

 
' Interface pointers
Private IAudioEndpointVolumeInterface As IAudioEndpointVolumeVtbl
Private m_pEndptVol As Long

Private IMMDeviceEnumeratorInterface As IMMDeviceEnumeratorVtbl
Private pEnumerator As Long

Private IMMDeviceInterface As IMMDeviceVtbl
Private pDevice As Long

' Event context
Private pguidEventContext As UUID

' HRESULT
Private hr As Long

 
Private Sub Class_Initialize()
  
  ' Instance
  Dim vtbl                    As Long
  Dim MMDeviceEnumerator      As UUID
  Dim IMMDeviceEnumerator     As UUID
  Dim IAudioEndpointVolume    As UUID

  ' Convert clsid, riid
  IIDFromString StrPtr(CLSID_MyContext), pguidEventContext
  IIDFromString StrPtr(CLSID_MMDeviceEnumerator), MMDeviceEnumerator
  IIDFromString StrPtr(IID_IMMDeviceEnumerator), IMMDeviceEnumerator
  IIDFromString StrPtr(IID_IAudioEndpointVolume), IAudioEndpointVolume
  
  ' Get enumerator for audio endpoint devices.
  hr = CoCreateInstance(MMDeviceEnumerator, 0, CLSCTX_INPROC_SERVER, IMMDeviceEnumerator, pEnumerator)
  
  If hr <> 0 Then
    Exit Sub
  End If
  
  ' vtbl
  RtlMoveMemory vtbl, ByVal pEnumerator, 4
  RtlMoveMemory IMMDeviceEnumeratorInterface, ByVal vtbl, Len(IMMDeviceEnumeratorInterface)
  ' Get default audio-rendering device.
  hr = CallPointer(IMMDeviceEnumeratorInterface.GetDefaultAudioEndpoint, pEnumerator, eRender, eConsole, VarPtr(pDevice))
 
  If hr <> 0 Then
    Exit Sub
  End If
  
  ' vtbl
  RtlMoveMemory vtbl, ByVal pDevice, 4
  RtlMoveMemory IMMDeviceInterface, ByVal vtbl, Len(IMMDeviceInterface)
  ' Get IAudioEndpointVolume interface
  hr = CallPointer(IMMDeviceInterface.Activate, pDevice, VarPtr(IAudioEndpointVolume), CLSCTX_ALL, 0, VarPtr(m_pEndptVol))
  
  If hr <> 0 Then
    Exit Sub
  End If
 
  ' vtbl
  RtlMoveMemory vtbl, ByVal m_pEndptVol, 4
  RtlMoveMemory IAudioEndpointVolumeInterface, ByVal vtbl, Len(IAudioEndpointVolumeInterface)
  
End Sub

Private Sub Class_Terminate()
  ' Cleanup
  If pDevice <> 0 Then
    hr = CallPointer(IMMDeviceInterface.Release, pDevice)
  End If
  If m_pEndptVol <> 0 Then
    hr = CallPointer(IAudioEndpointVolumeInterface.Release, m_pEndptVol)
  End If
End Sub

Public Function GetChannelCount() As Long
  ' The GetChannelCount method gets a count of the channels in the audio stream that enters or leaves the audio endpoint device.
  Dim pnChannelCount As Long
  hr = CallPointer(IAudioEndpointVolumeInterface.GetChannelCount, m_pEndptVol, VarPtr(pnChannelCount))
  GetChannelCount = pnChannelCount
End Function

Public Function GetChannelVolumeLevel(ByVal nChannel As Long) As Single
  ' The GetChannelVolumeLevel method gets the volume level, in decibels, of the specified channel in the audio stream that enters
  ' or leaves the audio endpoint device.
  Dim pfLevelDB As Single
  hr = CallPointer(IAudioEndpointVolumeInterface.GetChannelVolumeLevel, m_pEndptVol, nChannel, VarPtr(pfLevelDB))
  GetChannelVolumeLevel = pfLevelDB
End Function

Public Function GetVolumeValue(Value As String) As Integer
Dim SplitValue

If InStr(1, Value, ",") > 0 Then
    SplitValue = Split(Value, ",")
    GetVolumeValue = SplitValue(1) * 10
Else
    GetVolumeValue = Value * 10
End If
End Function
Public Function GetChannelVolumeLevelScalar(ByVal nChannel As Long) As Single
  ' The GetChannelVolumeLevelScalar method gets the normalized, audio-tapered volume level of the specified channel of the audio
  ' stream that enters or leaves the audio endpoint device.
  Dim pfLevel As Single
  hr = CallPointer(IAudioEndpointVolumeInterface.GetChannelVolumeLevelScalar, m_pEndptVol, nChannel, VarPtr(pfLevel))
  GetChannelVolumeLevelScalar = pfLevel
End Function

Public Function GetMasterVolumeLevel() As Single
  ' The GetMasterVolumeLevel method gets the master volume level, in decibels, of the audio stream that enters or leaves the audio
  ' endpoint device.
  Dim pfLevelDB As Single
  hr = CallPointer(IAudioEndpointVolumeInterface.GetMasterVolumeLevel, m_pEndptVol, VarPtr(pfLevelDB))
  GetMasterVolumeLevel = pfLevelDB
End Function

Public Function GetMasterVolumeLevelScalar() As Single
  ' The GetMasterVolumeLevelScalar method gets the master volume level of the audio stream that enters or leaves the audio endpoint
  ' device. The volume level is expressed as a normalized, audio-tapered value in the range from 0.0 to 1.0.
  Dim pfLevel As Single
  hr = CallPointer(IAudioEndpointVolumeInterface.GetMasterVolumeLevelScalar, m_pEndptVol, VarPtr(pfLevel))
  GetMasterVolumeLevelScalar = pfLevel
End Function

Public Function GetMute() As Long
  ' The GetMute method gets the muting state of the audio stream that enters or leaves the audio endpoint device.
  Dim pbMute As Long
  hr = CallPointer(IAudioEndpointVolumeInterface.GetMute, m_pEndptVol, VarPtr(pbMute))
  GetMute = pbMute
End Function

Public Sub GetVolumeRange(ByRef pfLevelMinDB As Single, ByRef pfLevelMaxDB As Single, ByRef pfVolumeIncrementDB As Single) '[out]
  ' The GetVolumeRange method gets the volume range, in decibels, of the audio stream that enters or leaves the audio endpoint device.
  hr = CallPointer(IAudioEndpointVolumeInterface.GetVolumeRange, m_pEndptVol, VarPtr(pfLevelMinDB), VarPtr(pfLevelMaxDB), VarPtr(pfVolumeIncrementDB))
End Sub

Public Sub GetVolumeStepInfo(ByRef pnStep As Long, ByRef pnStepCount As Long) '[out]
  ' The GetVolumeStepInfo method gets information about the current step in the volume range.
  hr = CallPointer(IAudioEndpointVolumeInterface.GetVolumeStepInfo, m_pEndptVol, VarPtr(pnStep), VarPtr(pnStepCount))
End Sub

Public Function QueryHardwareSupport() As Long
  ' The QueryHardwareSupport method queries the audio endpoint device for its hardware-supported functions.
  Dim pdwHardwareSupportMask As Long
  hr = CallPointer(IAudioEndpointVolumeInterface.QueryHardwareSupport, m_pEndptVol, VarPtr(pdwHardwareSupportMask))
  QueryHardwareSupport = pdwHardwareSupportMask
End Function

Public Sub RegisterControlChangeNotify(ByVal ptr As Long)
  ' The RegisterControlChangeNotify method registers a client's notification callback interface.
  hr = CallPointer(IAudioEndpointVolumeInterface.RegisterControlChangeNotify, m_pEndptVol, ptr)
End Sub

Public Sub SetChannelVolumeLevel(ByVal nChannel As Long, ByVal fLevelDB As Single)
  ' The SetChannelVolumeLevel method sets the volume level, in decibels, of the specified channel of the audio stream that enters
  ' or leaves the audio endpoint device.
  hr = CallPointer(IAudioEndpointVolumeInterface.SetChannelVolumeLevel, m_pEndptVol, nChannel, FloatToLongForCallee(fLevelDB), VarPtr(pguidEventContext))
End Sub

Public Sub SetChannelVolumeLevelScalar(ByVal nChannel As Long, ByVal fLevel As Single)
  ' The SetChannelVolumeLevelScalar method sets the normalized, audio-tapered volume level of the specified channel in the audio
  ' stream that enters or leaves the audio endpoint device.
  hr = CallPointer(IAudioEndpointVolumeInterface.SetChannelVolumeLevelScalar, m_pEndptVol, nChannel, FloatToLongForCallee(fLevel), VarPtr(pguidEventContext))
End Sub

Public Sub SetMasterVolumeLevel(ByVal fLevelDB As Single)
  ' The SetMasterVolumeLevel method sets the master volume level, in decibels, of the audio stream that enters or leaves the audio
  ' endpoint device.
  hr = CallPointer(IAudioEndpointVolumeInterface.SetMasterVolumeLevel, m_pEndptVol, FloatToLongForCallee(fLevelDB), VarPtr(pguidEventContext))
End Sub

Public Sub SetMasterVolumeLevelScalar(ByVal fLevel As Single)
  ' The SetMasterVolumeLevelScalar method sets the master volume level of the audio stream that enters or leaves the audio endpoint
  ' device. The volume level is expressed as a normalized, audio-tapered value in the range from 0.0 to 1.0.
  hr = CallPointer(IAudioEndpointVolumeInterface.SetMasterVolumeLevelScalar, m_pEndptVol, FloatToLongForCallee(fLevel), VarPtr(pguidEventContext))
End Sub

Public Sub SetMute(ByVal bMute As Long)
  ' The SetMute method sets the muting state of the audio stream that enters or leaves the audio endpoint device.
  hr = CallPointer(IAudioEndpointVolumeInterface.SetMute, m_pEndptVol, bMute, VarPtr(pguidEventContext))
End Sub

Public Sub UnregisterControlChangeNotify(ByVal ptr As Long)
  ' The UnregisterControlChangeNotify method deletes the registration of a client's notification callback interface that the client
  ' registered in a previous call to the IAudioEndpointVolume::RegisterControlChangeNotify method.
  hr = CallPointer(IAudioEndpointVolumeInterface.UnregisterControlChangeNotify, m_pEndptVol, ptr)
End Sub

Public Sub VolumeStepDown()
  ' The VolumeStepDown method decrements, by one step, the volume level of the audio stream that enters or leaves the audio endpoint device.
  hr = CallPointer(IAudioEndpointVolumeInterface.VolumeStepDown, m_pEndptVol, VarPtr(pguidEventContext))
End Sub

Public Sub VolumeStepUp()
  ' The VolumeStepUp method increments, by one step, the volume level of the audio stream that enters or leaves the audio endpoint device.
  hr = CallPointer(IAudioEndpointVolumeInterface.VolumeStepUp, m_pEndptVol, VarPtr(pguidEventContext))
End Sub

Private Function FloatToLongForCallee(ByVal Value As Single) As Long
  ' Converts a single floating point to a long value. The callee  will convert it back to the real floating value for consumption.
  Dim dwFloat As Long
  RtlMoveMemory dwFloat, Value, 4
  FloatToLongForCallee = dwFloat
End Function

Private Function CallPointer(ByVal fnc As Long, ParamArray params()) As Long

  Dim btASM(&HEC00& - 1) As Byte
  Dim pASM As Long
  Dim I As Integer
      
  pASM = VarPtr(btASM(0))
      
  AddByte pASM, &H58 '   POP   EAX
  AddByte pASM, &H59 '   POP   ECX
  AddByte pASM, &H59 '   POP   ECX
  AddByte pASM, &H59 '   POP   ECX
  AddByte pASM, &H59 '   POP   ECX
  AddByte pASM, &H50 '   PUSH  EAX
      
  For I = UBound(params) To 0 Step -1
            AddPush pASM, CLng(params(I))
  Next
      
  AddCall pASM, fnc
  AddByte pASM, &HC3
      
  CallPointer = CallWindowProcA(VarPtr(btASM(0)), 0, 0, 0, 0)
End Function

Private Sub AddPush(pASM As Long, lng As Long)
  AddByte pASM, &H68
  AddLong pASM, lng
End Sub
    
Private Sub AddCall(pASM As Long, addr As Long)
  AddByte pASM, &HE8
  AddLong pASM, addr - pASM - 4
End Sub
    
Private Sub AddLong(pASM As Long, lng As Long)
  RtlMoveMemory ByVal pASM, lng, 4
  pASM = pASM + 4
End Sub
    
Private Sub AddByte(pASM As Long, bt As Byte)
  RtlMoveMemory ByVal pASM, bt, 1
  pASM = pASM + 1
End Sub

