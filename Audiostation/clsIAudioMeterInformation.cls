VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub OleInitialize Lib "ole32.dll" (pvReserved As Any)
Private Declare Sub OleUninitialize Lib "ole32.dll" ()
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As String, pclsid As UUID) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As String, lpiid As UUID) As Long
Private Declare Function CoCreateInstance Lib "ole32.dll" (rclsid As UUID, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As UUID, ppv As Any) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PutMem2 Lib "msvbvm60" (ByVal pWORDDst As Long, ByVal NewValue As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal pDWORDDst As Long, ByVal NewValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal pDWORDSrc As Long, ByVal pDWORDDst As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Const GMEM_FIXED As Long = &H0
Private Const asmPUSH_imm32 As Byte = &H68
Private Const asmRET_imm16 As Byte = &HC2
Private Const asmCALL_rel32 As Byte = &HE8

Private Const CLSCTX_INPROC_SERVER                      As Long = &H1
Private Const CLSCTX_ALL                                As Long = &H0

Private Const UUIDOF_MMDeviceEnumerator As String = "{bcde0395-e52f-467c-8e3d-c4579291692e}"
Private Const UUIDOF_IMMDeviceEnumerator As String = "{a95664d2-9614-4f35-a746-de8db63617e6}"
Private Const UUIDOF_IAudioMeterInformation As String = "{C02216F6-8C67-4B5B-9D00-D008E73E0064}"

Private Type UUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Enum IUnknown_Exports
    [QueryInterface] = 0
    [AddRef] = 1
    [Release] = 2
End Enum

Private Enum ERole
    [eConsole]
    [eMultimedia]
    [eCommunications]
End Enum

Private Enum EDataFlow
    [eRender]
    [eCapture]
    [eAll]
End Enum

Private Enum IMMDeviceEnumerator_Exports
    [EnumAudioEndpoints] = 3
    [GetDefaultAudioEndpoint] = 4               'params=3
    [GetDevice] = 5
    [RegisterEndpointNotificationCallback] = 6
    [UnregisterEndpointNotificationCallback]
End Enum

Private Enum IAudioMeterInformation_Exports
    [GetPeakValue] = 3                          'params=1
    [GetMeteringChannelCount] = 4               'params=1
    [GetChannelsPeakValues] = 5                 'params=2
    [QueryHardwareSupport] = 6
End Enum

Private Enum IMMDevice_Exports
    [Activate] = 3                              'params=4
    [OpenPropertyStore] = 4
    [GetId] = 5
    [GetState] = 6
End Enum

Private c_lngObjDevEnumerator                           As Long
Private c_lngObjIMMDevice                               As Long
Private c_lngObjAudioMeterInformation                   As Long
Private c_blnInitialized                                As Boolean

'    // Get enumerator for audio endpoint devices.
'    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator),
'                          NULL, CLSCTX_INPROC_SERVER,
'                          __uuidof(IMMDeviceEnumerator),
'                          (void**)&pEnumerator);


'    // Get peak meter for default audio-rendering device.
'    hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);


'    hr = pDevice->Activate(__uuidof(IAudioMeterInformation),
'                           CLSCTX_ALL, NULL, (void**)&pMeterInfo);


Private Sub Class_Initialize()
    Dim uuidMMDeviceEnumerator As UUID
    Dim uuidIMMDeviceEnumerator As UUID
    Dim uuidIAudioMeterInformation As UUID
    
    If IsVista Then
        Call IIDFromString(StrConv(UUIDOF_MMDeviceEnumerator, vbUnicode), uuidMMDeviceEnumerator)
        Call IIDFromString(StrConv(UUIDOF_IMMDeviceEnumerator, vbUnicode), uuidIMMDeviceEnumerator)
        Call IIDFromString(StrConv(UUIDOF_IAudioMeterInformation, vbUnicode), uuidIAudioMeterInformation)
        
        Call CoCreateInstance(uuidMMDeviceEnumerator, 0, CLSCTX_INPROC_SERVER, uuidIMMDeviceEnumerator, c_lngObjDevEnumerator)
        If Not (c_lngObjDevEnumerator = 0) Then
            Call CallInterface(c_lngObjDevEnumerator, [GetDefaultAudioEndpoint], 3, [eRender], [eConsole], VarPtr(c_lngObjIMMDevice))
            If Not (c_lngObjIMMDevice = 0) Then
                Call CallInterface(c_lngObjIMMDevice, [Activate], 4, VarPtr(uuidIAudioMeterInformation), CLSCTX_ALL, 0, VarPtr(c_lngObjAudioMeterInformation))
                c_blnInitialized = Not (c_lngObjAudioMeterInformation = 0)
            End If
        End If
    End If
End Sub

Private Sub Class_Terminate()
    If IsVista Then
        If Not (c_lngObjDevEnumerator = 0) Then Call CallInterface(c_lngObjDevEnumerator, [Release], 0)
        If Not (c_lngObjIMMDevice = 0) Then Call CallInterface(c_lngObjIMMDevice, [Release], 0)
        If Not (c_lngObjAudioMeterInformation = 0) Then Call CallInterface(c_lngObjAudioMeterInformation, [Release], 0)
    End If
End Sub

Private Function IsVista() As Boolean
    IsVista = (((GetVersion() And &HFFFF&) Mod 256) >= 6)
End Function

Public Function GetPeak() As Single
    If IsVista Then
        If c_blnInitialized Then
            Call CallInterface(c_lngObjAudioMeterInformation, [GetPeakValue], 1, VarPtr(GetPeak))
        End If
    End If
End Function

Public Function GetChannelPeak(ByVal lngChannel As Long) As Single
    Dim lngChannels As Long
    Dim sngChannels() As Single
    
    If IsVista Then
        If c_blnInitialized Then
            Call CallInterface(c_lngObjAudioMeterInformation, [GetMeteringChannelCount], 1, VarPtr(lngChannels))
            ReDim sngChannels(lngChannels)
            Call CallInterface(c_lngObjAudioMeterInformation, [GetChannelsPeakValues], 2, lngChannels, VarPtr(sngChannels(0)))
            GetChannelPeak = sngChannels(lngChannel)
        End If
    End If
End Function

Private Function CallInterface(ByVal pInterface As Long, ByVal Member As Long, ByVal ParamsCount As Long, Optional ByVal p1 As Long = 0, Optional ByVal p2 As Long = 0, Optional ByVal p3 As Long = 0, Optional ByVal p4 As Long = 0, Optional ByVal p5 As Long = 0, Optional ByVal p6 As Long = 0, Optional ByVal p7 As Long = 0, Optional ByVal p8 As Long = 0, Optional ByVal p9 As Long = 0, Optional ByVal p10 As Long = 0) As Long
  Dim I As Long, t As Long
  Dim hGlobal As Long, hGlobalOffset As Long
  
  If ParamsCount < 0 Then err.Raise 5 'invalid call
  If pInterface = 0 Then err.Raise 5
  
  hGlobal = GlobalAlloc(GMEM_FIXED, 5 * ParamsCount + 5 + 5 + 3 + 1)
  If hGlobal = 0 Then err.Raise 7 'insuff. memory
  hGlobalOffset = hGlobal
  
  If ParamsCount > 0 Then
    t = VarPtr(p1)
    For I = ParamsCount - 1 To 0 Step -1
      Call PutMem2(hGlobalOffset, asmPUSH_imm32)
      hGlobalOffset = hGlobalOffset + 1
      Call GetMem4(t + I * 4, hGlobalOffset)
      hGlobalOffset = hGlobalOffset + 4
    Next
  End If
  
  Call PutMem2(hGlobalOffset, asmPUSH_imm32)
  hGlobalOffset = hGlobalOffset + 1
  Call PutMem4(hGlobalOffset, pInterface)
  hGlobalOffset = hGlobalOffset + 4
  
  Call PutMem2(hGlobalOffset, asmCALL_rel32)
  hGlobalOffset = hGlobalOffset + 1
  Call GetMem4(pInterface, VarPtr(t))     'дереференс: находим положение vTable
  Call GetMem4(t + Member * 4, VarPtr(t)) 'смещение по vTable, после чего дереференс оного
  Call PutMem4(hGlobalOffset, t - hGlobalOffset - 4)
  hGlobalOffset = hGlobalOffset + 4
   
  Call PutMem4(hGlobalOffset, &H10C2&)        'ret 0x0010
  
  CallInterface = CallWindowProc(hGlobal, 0, 0, 0, 0)
  
  Call GlobalFree(hGlobal)
End Function

