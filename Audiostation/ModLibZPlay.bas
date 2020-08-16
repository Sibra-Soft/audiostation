Attribute VB_Name = "ModLibZPlay"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : ModLibZPlay
'    Project    : Audiostation
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Public Enum TSettingID
    sidWaveBufferSize = 1
    sidAccurateLength = 2
    sidAccurateSeek = 3
    sidSamplerate = 4
    sidChannelNumber = 5
    sidBitPerSample = 6
    sidBigEndian = 7
End Enum

Public Enum TStreamFormat
    sfUnknown = 0
    sfMp3 = 1
    sfOgg = 2
    sfWav = 3
    sfPCM = 4
    sfFLAC = 5
    sfFLACOgg = 6
    sfAC3 = 7
    sfAacADTS = 8
    sfWaveIn = 9
    sfAutodetect = 1000
End Enum

Public Enum TFFTWindow
    fwRectangular = 1
    fwHamming
    fwHann
    fwCosine
    fwLanczos
    fwBartlett
    fwTriangular
    fwGauss
    fwBartlettHann
    fwBlackman
    fwNuttall
    fwBlackmanHarris
    fwBlackmanNuttall
    fwFlatTop
End Enum
    
Public Enum TTimeFormat
    tfMillisecond = 1
    tfSecond = 2
    tfHMS = 4
    tfSamples = 8
End Enum

Public Enum TSeekMethod
    smFromBeginning = 1
    smFromEnd = 2
    smFromCurrentForward = 4
    smFromCurrentBackward = 8
End Enum

Public Enum TID3Version
    id3Version1 = 1
    id3Version2 = 2
End Enum

Public Enum TFFTGraphHorizontalScale
    gsLogarithmic = 0
    gsLinear = 1
End Enum

Public Enum TFFTGraphParamID
    gpFFTPoints = 1
    gpGraphType
    gpWindow
    gpHorizontalScale
    gpSubgrid
    gpTransparency
    gpFrequencyScaleVisible
    gpDecibelScaleVisible
    gpFrequencyGridVisible
    gpDecibelGridVisible
    gpBgBitmapVisible
    gpBgBitmapHandle
    gpColor1
    gpColor2
    gpColor3
    gpColor4
    gpColor5
    gpColor6
    gpColor7
    gpColor8
    gpColor9
    gpColor10
    gpColor11
    gpColor12
    gpColor13
    gpColor14
    gpColor15
    gpColor16
End Enum

Public Enum TFFTGraphType
    gtLinesLeftOnTop = 0
    gtLinesRightOnTop
    gtAreaLeftOnTop
    gtAreaRightOnTop
    gtBarsLeftOnTop
    gtBarsRightOnTop
    gtSpectrum
End Enum

Public Enum TBPMDetectionMethod
    dmPeaks = 0
    dmAutoCorrelation
End Enum

Public Enum TFFTGraphSize
    FFTGraphMinWidth = 100
    FFTGraphMinHeight = 60
End Enum

Public Enum TWaveOutMapper
    WaveOutWaveMapper = &HFFFFFFFF '4294967295
End Enum

Public Enum TWaveInMapper
    WaveInWaveMapper = &HFFFFFFFF '4294967295
End Enum

Public Enum TCallbackMessage
    MsgStopAsync = 1
    MsgPlayAsync = 2
    MsgEnterLoopAsync = 4
    MsgExitLoopAsync = 8
    MsgEnterVolumeSlideAsync = 16
    MsgExitVolumeSlideAsync = 32
    MsgStreamBufferDoneAsync = 64
    MsgStreamNeedMoreDataAsync = 128
    MsgNextSongAsync = 256
    MsgStop = 65536
    MsgPlay = 131072
    MsgEnterLoop = 262144
    MsgExitLoop = 524288
    MsgEnterVolumeSlide = 1048576
    MsgExitVolumeSlide = 2097152
    MsgStreamBufferDone = 4194304
    MsgStreamNeedMoreData = 8388608
    MsgNextSong = 16777216
    MsgWaveBuffer = 33554432
End Enum

Public Enum TWaveOutFormat
    format_invalid = 0
    format_11khz_8bit_mono = 1
    format_11khz_8bit_stereo = 2
    format_11khz_16bit_mono = 4
    format_11khz_16bit_stereo = 8
    format_22khz_8bit_mono = 16
    format_22khz_8bit_stereo = 32
    format_22khz_16bit_mono = 64
    format_22khz_16bit_stereo = 128
    format_44khz_8bit_mono = 256
    format_44khz_8bit_stereo = 512
    format_44khz_16bit_mono = 1024
    format_44khz_16bit_stereo = 2048
End Enum

Public Enum TWaveOutFunctionality
    supportPitchControl = 1
    supportPlaybackRateControl = 2
    supportVolumeControl = 4
    supportSeparateLeftRightVolume = 8
    supportSync = 16
    supportSampleAccuratePosition = 32
    supportDirectSound = 6
End Enum
' ----==== libZplay Public Enums ====----

' ----==== libZplay Public Type ====----
Public Type TStreamHMSTime
    hour As Long
    minute As Long
    second As Long
    millisecond As Long
End Type

Public Type TStreamTime
    sec As Long
    ms As Long
    samples As Long
    hms As TStreamHMSTime
End Type

Public Type TStreamInfo
    SamplingRate As Long
    ChannelNumber As Long
    VBR As Long
    BitRate As Long
    Length As TStreamTime
    Description As String
End Type

Public Type TWaveOutInfo
    ManufacturerID As Long
    ProductID As Long
    DriverVersion As Long
    Formats As Long
    Channels As Long
    Support As Long
    ProductName As String
End Type

Public Type TWaveInInfo
    ManufacturerID As Long
    ProductID As Long
    DriverVersion As Long
    Formats As Long
    Channels As Long
    ProductName As String
End Type

Public Type TStreamLoadInfo
    NumberOfBuffers As Long
    NumberOfBytes As Long
End Type

Public Type TEchoEffect
    nLeftDelay As Long
    nLeftSrcVolume As Long
    nLeftEchoVolume As Long
    nRightDelay As Long
    nRightSrcVolume As Long
    nRightEchoVolume As Long
End Type

Public Type TStreamStatus
    fPlay As Long
    fPause As Long
    fEcho As Long
    fEqualizer As Long
    fVocalCut As Long
    fSideCut As Long
    fChannelMix As Long
    fSlideVolume As Long
    nLoop As Long
    fReverse As Long
    nSongIndex As Long
    nSongsInQueue As Long
End Type

Public Type TID3Info
    Title As String
    Artist As String
    Album As String
    Year As String
    Comment As String
    Track As String
    Genre As String
End Type

Public Type TID3Picture
    PicturePresent As Boolean
    PictureType As Long
    Description As String
    Bitmap As StdPicture
    BitStream As IUnknown
End Type

Public Type TID3InfoEx
    Title As String
    Artist As String
    Album As String
    Year As String
    Comment As String
    Track As String
    Genre As String
    AlbumArtist As String
    Composer As String
    OriginalArtist As String
    Copyright As String
    url As String
    Encoder As String
    Publisher As String
    BPM As Long
    Picture As TID3Picture
End Type
' ----==== libZplay Public Type ====----

' ----==== libZplay Private Type Internal ====----
Private Type TStreamInfo_Internal
    SamplingRate As Long
    ChannelNumber As Long
    VBR As Long
    BitRate As Long
    Length As TStreamTime
    Description As Long
End Type

Private Type TWaveOutInfo_Internal
    ManufacturerID As Long
    ProductID As Long
    DriverVersion As Long
    Formats As Long
    Channels As Long
    Support As Long
    ProductName As Long
End Type

Private Type TWaveInInfo_Internal
    ManufacturerID As Long
    ProductID As Long
    DriverVersion As Long
    Formats As Long
    Channels As Long
    ProductName As Long
End Type

Private Type TID3Info_Internal
    Title As Long
    Artist As Long
    Album As Long
    Year As Long
    Comment As Long
    Track As Long
    Genre As Long
End Type

Private Type TID3InfoEx_Internal
     Title As Long
     Artist As Long
     Album As Long
     Year As Long
     Comment As Long
     Track As Long
     Genre As Long
     AlbumArtist As Long
     Composer As Long
     OriginalArtist As Long
     Copyright As Long
     url As Long
     Encoder As Long
     Publisher As Long
     BPM As Long
     PicturePresent As Long
     CanDrawPicture As Long
     MIMEType As Long
     PictureType As Long
     Description As Long
     PictureData As Long
     PictureDataSize As Long
     hBitmap As Long
     Width As Long
     Height As Long
     Dummy(65) As Long
     Reserved As Long
End Type
' ----==== libZplay Private Type Internal ====----

' ----==== libZplay Declare ====----
Private Declare Function zplay_CreateZPlay Lib "libzplay" () As Long

Private Declare Function zplay_DestroyZPlay Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_SetSettings Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nSettingID As Long, _
                         ByVal Value As Long) As Long

Private Declare Function zplay_GetSettings Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nSettingID As Long) As Long

Private Declare Function zplay_GetError Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_GetErrorW Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_GetVersion Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_GetFileFormat Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As String) As Long

Private Declare Function zplay_GetFileFormatW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As Long) As Long

Private Declare Function zplay_OpenFile Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As String, _
                         ByVal nFormat As Long) As Long

Private Declare Function zplay_OpenFileW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As Long, _
                         ByVal nFormat As Long) As Long

Private Declare Function zplay_AddFile Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As String, _
                         ByVal nFormat As Long) As Long

Private Declare Function zplay_AddFileW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As Long, _
                         ByVal nFormat As Long) As Long

Private Declare Function zplay_OpenStream Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fBuffered As Long, _
                         ByVal fManaged As Long, _
                         ByRef sMemStream() As Byte, _
                         ByVal nStreamSize As Long, _
                         ByVal nFormat As Long) As Long

Private Declare Function zplay_PushDataToStream Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByRef sMemNewData() As Byte, _
                         ByVal nNewDataize As Long) As Long

Private Declare Function zplay_Close Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_Play Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_Stop Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_Pause Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_Resume Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_IsStreamDataFree Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByRef sMemNewData() As Byte) As Long

Private Declare Sub zplay_GetDynamicStreamLoad Lib "libzplay" ( _
                    ByVal objptr As Long, _
                    ByRef pStreamLoadInfo As TStreamLoadInfo)

Private Declare Sub zplay_GetPosition Lib "libzplay" ( _
                    ByVal objptr As Long, _
                    ByRef pTime As TStreamTime)

Private Declare Function zplay_PlayLoop Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fFormatStartTime As Long, _
                         ByRef pStartTime As TStreamTime, _
                         ByVal fFormatEndTime As Long, _
                         ByRef pEndTime As TStreamTime, _
                         ByVal nNumOfCycles As Long, _
                         ByVal fContinuePlaying As Long) As Long

Private Declare Function zplay_Seek Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fFormat As TTimeFormat, _
                         ByRef pTime As TStreamTime, _
                         ByVal nMoveMethod As TSeekMethod) As Long

Private Declare Function zplay_ReverseMode Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fEnable As Long) As Long

Private Declare Function zplay_SetMasterVolume Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nLeftVolume As Long, _
                         ByVal nRightVolume As Long) As Long

Private Declare Function zplay_SetPlayerVolume Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nLeftVolume As Long, _
                         ByVal nRightVolume As Long) As Long

Private Declare Sub zplay_GetMasterVolume Lib "libzplay" ( _
                    ByVal objptr As Long, _
                    ByRef nLeftVolume As Long, _
                    ByRef nRightVolume As Long)

Private Declare Sub zplay_GetPlayerVolume Lib "libzplay" ( _
                    ByVal objptr As Long, _
                    ByRef nLeftVolume As Long, _
                    ByRef nRightVolume As Long)

Private Declare Function zplay_GetBitrate Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fAverage As Long) As Long

Private Declare Sub zplay_GetStatus Lib "libzplay" ( _
                    ByVal objptr As Long, _
                    ByRef pStatus As TStreamStatus)

Private Declare Function zplay_MixChannels Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fEnable As Long, _
                         ByVal nLeftPercent As Long, _
                         ByVal nRightPercent As Long) As Long

Private Declare Sub zplay_GetVUData Lib "libzplay" ( _
                    ByVal objptr As Long, _
                    ByRef pnLeftChannel As Long, _
                    ByRef pnRightChannel As Long)

Private Declare Function zplay_SlideVolume Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fFormatStart As TTimeFormat, _
                         ByRef pTimeStart As TStreamTime, _
                         ByVal nStartVolumeLeft As Long, _
                         ByVal nStartVolumeRight As Long, _
                         ByVal fFormatEnd As TTimeFormat, _
                         ByRef pTimeEnd As TStreamTime, _
                         ByVal nEndVolumeLeft As Long, _
                         ByVal nEndVolumeRight As Long) As Long

Private Declare Function zplay_EnableEqualizer Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fEnable As Long) As Long

Private Declare Function zplay_SetEqualizerPoints Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByRef pnFreqPoint() As Long, _
                         ByVal nNumOfPoints As Long) As Long

Private Declare Function zplay_GetEqualizerPoints Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByRef pnFreqPoint() As Long, _
                         ByVal nNumOfPoints As Long) As Long

Private Declare Function zplay_SetEqualizerParam Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nPreAmpGain As Long, _
                         ByRef pnBandGain() As Long, _
                         ByVal nNumberOfBands As Long) As Long

Private Declare Function zplay_GetEqualizerParam Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nPreAmpGain As Long, _
                         ByRef pnBandGain() As Long, _
                         ByVal nNumberOfBands As Long) As Long

Private Declare Function zplay_SetEqualizerPreampGain Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nGain As Long) As Long

Private Declare Function zplay_GetEqualizerPreampGain Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_SetEqualizerBandGain Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nBandIndex As Long, _
                         ByVal nGain As Long) As Long

Private Declare Function zplay_GetEqualizerBandGain Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nBandIndex As Long) As Long

Private Declare Function zplay_EnableEcho Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fEnable As Long) As Long

Private Declare Function zplay_StereoCut Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal fEnable As Long, _
                         ByVal fOutputCenter As Long, _
                         ByVal fBassToSides As Long) As Long

Private Declare Function zplay_SetEchoParam Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByRef pEchoEffect() As TEchoEffect, _
                         ByVal nNumberOfEffects As Long) As Long

Private Declare Function zplay_GetEchoParam Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByRef pEchoEffect() As TEchoEffect, _
                         ByVal nNumberOfEffects As Long) As Long

Private Declare Function zplay_GetFFTData Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nFFTPoints As Long, _
                         ByVal nFFTWindow As Long, _
                         ByRef pnHarmonicNumber As Long, _
                         ByRef pnHarmonicFreq() As Long, _
                         ByRef pnLeftAmplitude() As Long, _
                         ByRef pnRightAmplitude() As Long, _
                         ByRef pnLeftPhase() As Long, _
                         ByRef pnRightPhase() As Long) As Long

Private Declare Function zplay_SetRate Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nRate As Long) As Long

Private Declare Function zplay_GetRate Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_SetPitch Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nPitch As Long) As Long

Private Declare Function zplay_GetPitch Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_SetTempo Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nTempo As Long) As Long

Private Declare Function zplay_GetTempo Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_DrawFFTGraphOnHDC Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal hdc As Long, _
                         ByVal nX As Long, _
                         ByVal nY As Long, _
                         ByVal nWidth As Long, _
                         ByVal nHeight As Long) As Long

Private Declare Function zplay_DrawFFTGraphOnHWND Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal hwnd As Long, _
                         ByVal nX As Long, _
                         ByVal nY As Long, _
                         ByVal nWidth As Long, _
                         ByVal nHeight As Long) As Long

Private Declare Function zplay_SetFFTGraphParam Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nParamID As Long, _
                         ByVal nParamValue As Long) As Long

Private Declare Function zplay_GetFFTGraphParam Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nParamID As Long) As Long

Private Declare Function zplay_LoadID3W Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nId3Version As Long, _
                         ByRef pId3Info As TID3Info_Internal) As Long

Private Declare Function zplay_LoadFileID3W Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal pchFileName As Long, _
                         ByVal nFormat As Long, _
                         ByVal nId3Version As Long, _
                         ByRef pId3Info As TID3Info_Internal) As Long

Private Declare Function zplay_DetectBPM Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nMethod As Long) As Long

Private Declare Function zplay_DetectFileBPM Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As String, _
                         ByVal nFormat As Long, _
                         ByVal nMethod As Long) As Long

Private Declare Function zplay_DetectFileBPMW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As Long, _
                         ByVal nFormat As Long, _
                         ByVal nMethod As Long) As Long

Private Declare Function zplay_SetCallbackFunc Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal pCallbackFunc As Long, _
                         ByVal nMessage As TCallbackMessage, _
                         ByVal user_data As Long) As Long

Private Declare Function zplay_EnumerateWaveIn Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_GetWaveInInfoW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nIndex As Long, _
                         ByRef pWaveInInfo As TWaveInInfo_Internal) As Long

Private Declare Function zplay_SetWaveInDevice Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nIndex As Long) As Long

Private Declare Function zplay_EnumerateWaveOut Lib "libzplay" ( _
                         ByVal objptr As Long) As Long

Private Declare Function zplay_GetWaveOutInfoW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nIndex As Long, _
                         ByRef pWaveOutInfo As TWaveOutInfo_Internal) As Long

Private Declare Function zplay_SetWaveOutDevice Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal nIndex As Long) As Long

Private Declare Function zplay_GetStreamInfoW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByRef pInfo As TStreamInfo_Internal) As Long

Private Declare Function zplay_LoadID3ExW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByRef pInfo As TID3InfoEx_Internal, _
                         ByVal fDecodeEmbededPicture As Long) As Long

Private Declare Function zplay_LoadFileID3ExW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As Long, _
                         ByVal nFormat As Long, _
                         ByRef pInfo As TID3InfoEx_Internal, _
                         ByVal fDecodeEmbededPicture As Long) As Long

Private Declare Function zplay_SetWaveOutFileW Lib "libzplay" ( _
                         ByVal objptr As Long, _
                         ByVal sFileName As Long, _
                         ByVal nFormat As Long, _
                         ByVal fOutputToSoundcard As Long) As Long
' ----==== libZplay Declare ====----

' ----==== Kernel32 Declare ====----
Private Declare Sub CopyMemory Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)

Private Declare Function lstrlenW Lib "kernel32" ( _
                         ByVal lpString As Long) As Long
' ----==== Kernel32 Declare ====----

' ----==== OLE32 API Declarationen ====----
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" ( _
                         ByRef hGlobal As Any, _
                         ByVal fDeleteOnRelease As Long, _
                         ByRef ppstm As Any) As Long
                         

' ----==== OLEAUT32 Types ====----
Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

' ----==== OLEAUT32 API Declarations ====----
Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" ( _
                    ByRef lpPictDesc As PICTDESC, _
                    ByRef riid As IID, _
                    ByVal fOwn As Boolean, _
                    ByRef lplpvObj As Object)

' ----==== GDIPlus Types ====----
Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type

' ----==== GDI+ API Declarationen ====----
Private Declare Function GdiplusShutdown Lib "gdiplus" ( _
                         ByVal Token As Long) As Long
                         
Private Declare Function GdiplusStartup Lib "gdiplus" ( _
                         ByRef Token As Long, _
                         ByRef lpInput As GDIPlusStartupInput, _
                         ByRef lpOutput As GdiplusStartupOutput) As Long

Private Declare Function GdipLoadImageFromStream Lib "gdiplus" ( _
                         ByVal Stream As IUnknown, _
                         ByRef Image As Long) As Long
                         
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" ( _
                         ByVal Bitmap As Long, _
                         ByRef hBmReturn As Long, _
                         ByVal Background As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" ( _
                         ByVal Image As Long) As Long
                         
' ----==== Variablen ====----
Private GdipToken As Long
Private oPlayer As Long

Public Function TCallbackFunc(ByVal objptr As Long, ByVal user_data As Long, ByVal _
    msg As TCallbackMessage, ByVal param1 As Long, ByVal param2 As Long) As Long
    
    
    TCallbackFunc = 0
End Function

'#Region fsdf

Public Sub Initialize()
oPlayer = 0
oPlayer = zplay_CreateZPlay

If oPlayer = 0 Then

    Call MsgBox("Can't create libZPlay interface.", vbInformation Or vbOKOnly)

Else

    If GetVersion < 190 Then

        Call MsgBox("Need libZPlay.dll version 1.90 and above.", vbInformation _
            Or vbOKOnly)

    End If
End If
End Sub
'#End Region


Public Sub Terminate()

    If oPlayer <> 0 Then

        Call zplay_DestroyZPlay(oPlayer)

        oPlayer = 0

    End If

End Sub

'#Region "Version"
Public Function GetVersion() As Long

    If oPlayer <> 0 Then

        GetVersion = zplay_GetVersion(oPlayer)

    End If

End Function
'#End Region

'#Region "Error handling"
Public Function GetError() As String

    If oPlayer <> 0 Then

        GetError = PointerToString(zplay_GetErrorW(oPlayer))

    End If

End Function
' #End Region

'#Region "Open and close stream"
Public Function GetFileFormat(ByVal FileName As String) As TStreamFormat

    Dim pStr As Long

    If oPlayer <> 0 Then

        pStr = StrPtr(FileName)
        GetFileFormat = zplay_GetFileFormatW(oPlayer, pStr)

    End If

End Function

Public Function OpenFile(ByVal FileName As String, ByVal Format As TStreamFormat) _
    As Boolean

    Dim pStr As Long

    If oPlayer <> 0 Then

        pStr = StrPtr(FileName)
        OpenFile = CBool(zplay_OpenFileW(oPlayer, pStr, Format))

    End If

End Function

Public Function SetWaveOutFile(ByVal FileName As String, ByVal Format As _
    TStreamFormat, ByRef fOutputToSoundcard As Boolean) As Boolean

    Dim s As Long
    Dim pStr As Long

    If oPlayer <> 0 Then
        If fOutputToSoundcard Then

            s = 1

        End If

        pStr = StrPtr(FileName)
        SetWaveOutFile = CBool(zplay_SetWaveOutFileW(oPlayer, pStr, Format, s))

    End If

End Function

Public Function AddFile(ByVal FileName As String, ByVal Format As TStreamFormat) _
    As Boolean

    Dim pStr As Long

    If oPlayer <> 0 Then

        pStr = StrPtr(FileName)
        AddFile = CBool(zplay_AddFileW(oPlayer, pStr, Format))

    End If

End Function

Public Function OpenStream(ByVal Buffered As Boolean, ByVal Dynamic As Boolean, _
    ByRef MemStream() As Byte, ByVal StreamSize As Long, ByVal nFormat As _
    TStreamFormat) As Boolean

    Dim B As Long
    Dim m As Long

    If oPlayer <> 0 Then
        If Buffered Then

            B = 1

        End If

        If Dynamic Then

            m = 1

        End If

        OpenStream = CBool(zplay_OpenStream(oPlayer, B, m, MemStream, StreamSize, _
            nFormat))

    End If

End Function

Public Function PushDataToStream(ByRef MemNewData() As Byte, ByVal NewDatSize As _
    Long) As Boolean

    If oPlayer <> 0 Then

        PushDataToStream = CBool(zplay_PushDataToStream(oPlayer, MemNewData, _
            NewDatSize))

    End If

End Function

Public Function IsStreamDataFree(ByRef MemNewData() As Byte) As Boolean

    If oPlayer <> 0 Then

        IsStreamDataFree = CBool(zplay_IsStreamDataFree(oPlayer, MemNewData))

    End If

End Function

Public Function StreamClose() As Boolean

    If oPlayer <> 0 Then

        StreamClose = CBool(zplay_Close(oPlayer))

    End If

End Function
' #End Region

'#Region "Position and Seek"
Public Sub GetPosition(ByRef time As TStreamTime)
    If oPlayer <> 0 Then

        Call zplay_GetPosition(oPlayer, time)

    End If
End Sub

Public Function SeekPosition(ByVal TimeFormat As TTimeFormat, ByRef Position As _
    TStreamTime, ByVal MoveMethod As TSeekMethod) As Boolean

    If oPlayer <> 0 Then

        SeekPosition = CBool(zplay_Seek(oPlayer, TimeFormat, Position, MoveMethod))

    End If

End Function
'#End Region

'#Region "Play, Pause, Loop, Reverse"
Public Function ReverseMode(ByVal Enable As Boolean) As Boolean

    If oPlayer <> 0 Then
        If Enable Then

            ReverseMode = CBool(zplay_ReverseMode(oPlayer, 1))

        Else

            ReverseMode = CBool(zplay_ReverseMode(oPlayer, 0))

        End If
    End If

End Function

Public Function PlayLoop(ByVal TimeFormatStart As TTimeFormat, ByRef StartPosition _
    As TStreamTime, ByVal TimeFormatEnd As TTimeFormat, ByRef EndPosition As _
    TStreamTime, ByVal NumberOfCycles As Long, ByVal ContinuePlaying As Boolean) _
    As Boolean

    Dim continueplay As Long

    If oPlayer <> 0 Then
        If ContinuePlaying Then

            continueplay = 1

        Else

            continueplay = 0

        End If

        PlayLoop = CBool(zplay_PlayLoop(oPlayer, TimeFormatStart, _
            StartPosition, TimeFormatEnd, EndPosition, NumberOfCycles, _
            continueplay))

    End If

End Function

Public Function StartPlayback() As Boolean

    If oPlayer <> 0 Then

        StartPlayback = CBool(zplay_Play(oPlayer))

    End If

End Function

Public Function StopPlayback() As Boolean

    If oPlayer <> 0 Then

        StopPlayback = CBool(zplay_Stop(oPlayer))

    End If

End Function

Public Function PausePlayback() As Boolean

    If oPlayer <> 0 Then

        PausePlayback = CBool(zplay_Pause(oPlayer))

    End If

End Function

Public Function ResumePlayback() As Boolean

    If oPlayer <> 0 Then

        ResumePlayback = CBool(zplay_Resume(oPlayer))

    End If

End Function
' #End Region

'#Region "Equalizer"
Public Function SetEqualizerParam(ByVal PreAmpGain As Long, ByRef BandGain() As _
    Long, ByVal NumberOfBands As Long) As Boolean

    If oPlayer <> 0 Then

        SetEqualizerParam = CBool(zplay_SetEqualizerParam(oPlayer, PreAmpGain, _
            BandGain, NumberOfBands))

    End If

End Function
Public Function EnableEqualizer(ByVal Enable As Boolean) As Boolean

    If oPlayer <> 0 Then
        If Enable Then

            EnableEqualizer = CBool(zplay_EnableEqualizer(oPlayer, 1))

        Else

            EnableEqualizer = CBool(zplay_EnableEqualizer(oPlayer, 0))

        End If
    End If

End Function

Public Function SetEqualizerPreampGain(ByVal Gain As Long) As Boolean

    If oPlayer <> 0 Then

        SetEqualizerPreampGain = CBool(zplay_SetEqualizerPreampGain(oPlayer, Gain))

    End If

End Function

Public Function GetEqualizerPreampGain() As Long

    If oPlayer <> 0 Then

        GetEqualizerPreampGain = zplay_GetEqualizerPreampGain(oPlayer)

    End If

End Function

Public Function SetEqualizerBandGain(ByVal BandIndex As Long, ByVal Gain As Long) _
    As Boolean

    If oPlayer <> 0 Then

        SetEqualizerBandGain = CBool(zplay_SetEqualizerBandGain(oPlayer, _
            BandIndex, Gain))

    End If

End Function

Public Function GetEqualizerBandGain(ByVal BandIndex As Long) As Long

    If oPlayer <> 0 Then

        GetEqualizerBandGain = zplay_GetEqualizerBandGain(oPlayer, BandIndex)

    End If

End Function

Public Function SetEqualizerPoints(ByRef FreqPointArray() As Long, ByVal _
    NumberOfPoints As Long) As Boolean

    If oPlayer <> 0 Then

        SetEqualizerPoints = CBool(zplay_SetEqualizerPoints(oPlayer, _
            FreqPointArray, NumberOfPoints))

    End If

End Function

'#Region "Echo"
Public Function EnableEcho(ByVal Enable As Boolean) As Boolean

    If oPlayer <> 0 Then
        If Enable Then

            EnableEcho = CBool(zplay_EnableEcho(oPlayer, 1))

        Else

            EnableEcho = CBool(zplay_EnableEcho(oPlayer, 0))

        End If
    End If

End Function

Public Function SetEchoParam(ByRef EchoEffectArray() As TEchoEffect, ByVal _
    NumberOfEffects As Long) As Boolean

    If oPlayer <> 0 Then

        SetEchoParam = CBool(zplay_SetEchoParam(oPlayer, EchoEffectArray, _
            NumberOfEffects))

    End If

End Function
'#Region "Volume and Fade"
Public Function SetMasterVolume(ByVal LeftVolume As Long, ByVal RightVolume As _
    Long) As Boolean

    If oPlayer <> 0 Then

        SetMasterVolume = CBool(zplay_SetMasterVolume(oPlayer, LeftVolume, _
            RightVolume))

    End If

End Function

Public Function SetPlayerVolume(ByVal LeftVolume As Long, ByVal RightVolume As _
    Long) As Boolean

    If oPlayer <> 0 Then

        SetPlayerVolume = CBool(zplay_SetPlayerVolume(oPlayer, LeftVolume, _
            RightVolume))

    End If

End Function

Public Sub GetMasterVolume(ByRef LeftVolume As Long, ByRef RightVolume As Long)

    If oPlayer <> 0 Then

        Call zplay_GetMasterVolume(oPlayer, LeftVolume, RightVolume)

    End If

End Sub

Public Sub GetPlayerVolume(ByRef LeftVolume As Long, ByRef RightVolume As Long)

    If oPlayer <> 0 Then

        Call zplay_GetPlayerVolume(oPlayer, LeftVolume, RightVolume)

    End If

End Sub

Public Function SlideVolume(ByVal TimeFormatStart As TTimeFormat, ByRef TimeStart _
    As TStreamTime, ByVal StartVolumeLeft As Long, ByVal StartVolumeRight As Long, _
    ByVal TimeFormatEnd As TTimeFormat, ByRef TimeEnd As TStreamTime, ByVal _
    EndVolumeLeft As Long, ByVal EndVolumeRight As Long) As Boolean

    If oPlayer <> 0 Then

        SlideVolume = CBool(zplay_SlideVolume(oPlayer, TimeFormatStart, TimeStart, _
            StartVolumeLeft, StartVolumeRight, TimeFormatEnd, TimeEnd, _
            EndVolumeLeft, EndVolumeRight))

    End If

End Function
'#End Region

'#Region "Pitch, tempo, rate"
Public Function SetPitch(ByVal Pitch As Long) As Boolean

    If oPlayer <> 0 Then

        SetPitch = CBool(zplay_SetPitch(oPlayer, Pitch))

    End If

End Function

Public Function GetPitch() As Long

    If oPlayer <> 0 Then

        GetPitch = zplay_GetPitch(oPlayer)

    End If

End Function

Public Function SetRate(ByVal Rate As Long) As Boolean

    If oPlayer <> 0 Then

        SetRate = CBool(zplay_SetRate(oPlayer, Rate))

    End If

End Function

Public Function GetRate() As Long

    If oPlayer <> 0 Then

        GetRate = zplay_GetRate(oPlayer)

    End If

End Function

Public Function SetTempo(ByVal Tempo As Long) As Boolean

    If oPlayer <> 0 Then

        SetTempo = CBool(zplay_SetTempo(oPlayer, Tempo))

    End If

End Function

Public Function GetTempo() As Long

    If oPlayer <> 0 Then

        GetTempo = zplay_GetTempo(oPlayer)

    End If

End Function
'#End Region

'#Region "Bitrate"
Public Function GetBitrate(ByVal Average As Boolean) As Long

    If oPlayer <> 0 Then
        If Average Then

            GetBitrate = zplay_GetBitrate(oPlayer, 1)

        Else

            GetBitrate = zplay_GetBitrate(oPlayer, 0)

        End If
    End If

End Function
'#End Region

'#Region "ID3 Info"
Public Function LoadID3(ByVal Id3Version As TID3Version, ByRef Info As TID3Info) _
    As Boolean

    Dim tmp As TID3Info_Internal

    If oPlayer <> 0 Then
        If zplay_LoadID3W(oPlayer, Id3Version, tmp) = 1 Then

            Info.Album = PointerToString(tmp.Album)
            Info.Artist = PointerToString(tmp.Artist)
            Info.Comment = PointerToString(tmp.Comment)
            Info.Genre = PointerToString(tmp.Genre)
            Info.Title = PointerToString(tmp.Title)
            Info.Track = PointerToString(tmp.Track)
            Info.Year = PointerToString(tmp.Year)
            LoadID3 = True

        Else

            LoadID3 = False

        End If
    End If

End Function

Public Function LoadID3Ex(ByRef Info As TID3InfoEx, ByVal fDecodePicture As _
    Boolean) As Boolean

    Dim tmp As TID3InfoEx_Internal

    If oPlayer <> 0 Then
        If zplay_LoadID3ExW(oPlayer, tmp, 0) = 1 Then

            Info.Album = PointerToString(tmp.Album)
            Info.Artist = PointerToString(tmp.Artist)
            Info.Comment = PointerToString(tmp.Comment)
            Info.Genre = PointerToString(tmp.Genre)
            Info.Title = PointerToString(tmp.Title)
            Info.Track = PointerToString(tmp.Track)
            Info.Year = PointerToString(tmp.Year)
            Info.AlbumArtist = PointerToString(tmp.AlbumArtist)
            Info.Composer = PointerToString(tmp.Composer)
            Info.OriginalArtist = PointerToString(tmp.OriginalArtist)
            Info.Copyright = PointerToString(tmp.Copyright)
            Info.Encoder = PointerToString(tmp.Encoder)
            Info.Publisher = PointerToString(tmp.Publisher)
            Info.BPM = tmp.BPM
            Info.Picture.PicturePresent = False

            If fDecodePicture Then
                If tmp.PicturePresent = 1 Then

                    Dim stream_data() As Byte

                    ReDim stream_data(tmp.PictureDataSize)

                    CopyMemory stream_data(0), ByVal tmp.PictureData, _
                        tmp.PictureDataSize

                    If CreateStreamOnHGlobal(stream_data(0), False, _
                        Info.Picture.BitStream) = 0 Then

                        Set Info.Picture.Bitmap = StreamToPicture( _
                            Info.Picture.BitStream)

                        Set Info.Picture.BitStream = Nothing

                    End If

                    Info.Picture.PictureType = tmp.PictureType
                    Info.Picture.Description = PointerToString(tmp.Description)
                    Info.Picture.PicturePresent = True
                    LoadID3Ex = True

                Else

                    Set Info.Picture.Bitmap = LoadPicture("")

                    Info.Picture.PicturePresent = False

                End If
            End If

        Else

            LoadID3Ex = False

        End If
    End If

End Function

Public Function LoadFileID3(ByVal FileName As String, ByVal Format As _
    TStreamFormat, ByVal Id3Version As TID3Version, ByRef Info As TID3Info) As _
    Boolean

    Dim tmp As TID3Info_Internal
    Dim pStr As Long

    If oPlayer <> 0 Then

        pStr = StrPtr(FileName)

        If zplay_LoadFileID3W(oPlayer, pStr, Format, Id3Version, tmp) = 1 Then

            Info.Album = PointerToString(tmp.Album)
            Info.Artist = PointerToString(tmp.Artist)
            Info.Comment = PointerToString(tmp.Comment)
            Info.Genre = PointerToString(tmp.Genre)
            Info.Title = PointerToString(tmp.Title)
            Info.Track = PointerToString(tmp.Track)
            Info.Year = PointerToString(tmp.Year)
            LoadFileID3 = True

        Else

            LoadFileID3 = False

        End If
    End If

End Function

Public Function LoadFileID3Ex(ByVal FileName As String, ByVal Format As _
    TStreamFormat, ByRef Info As TID3InfoEx, ByVal fDecodePicture As Boolean) As _
    Boolean

    Dim tmp As TID3InfoEx_Internal
    Dim pStr As Long

    If oPlayer <> 0 Then

        pStr = StrPtr(FileName)

        If zplay_LoadFileID3ExW(oPlayer, pStr, Format, tmp, 0) = 1 Then

            Info.Album = PointerToString(tmp.Album)
            Info.Artist = PointerToString(tmp.Artist)
            Info.Comment = PointerToString(tmp.Comment)
            Info.Genre = PointerToString(tmp.Genre)
            Info.Title = PointerToString(tmp.Title)
            Info.Track = PointerToString(tmp.Track)
            Info.Year = PointerToString(tmp.Year)
            Info.AlbumArtist = PointerToString(tmp.AlbumArtist)
            Info.Composer = PointerToString(tmp.Composer)
            Info.OriginalArtist = PointerToString(tmp.OriginalArtist)
            Info.Copyright = PointerToString(tmp.Copyright)
            Info.Encoder = PointerToString(tmp.Encoder)
            Info.Publisher = PointerToString(tmp.Publisher)
            Info.BPM = tmp.BPM
            Info.Picture.PicturePresent = False

            If fDecodePicture Then
                If tmp.PicturePresent = 1 Then

                    Dim stream_data() As Byte

                    ReDim stream_data(tmp.PictureDataSize)

                    CopyMemory stream_data(0), ByVal tmp.PictureData, _
                        tmp.PictureDataSize

                    If CreateStreamOnHGlobal(stream_data(0), False, _
                        Info.Picture.BitStream) = 0 Then

                        Set Info.Picture.Bitmap = StreamToPicture( _
                            Info.Picture.BitStream)

                        Set Info.Picture.BitStream = Nothing

                    End If

                    Info.Picture.PictureType = tmp.PictureType
                    Info.Picture.Description = PointerToString(tmp.Description)
                    Info.Picture.PicturePresent = True
                    LoadFileID3Ex = True

                Else

                    Set Info.Picture.Bitmap = LoadPicture("")

                    Info.Picture.PicturePresent = False

                End If
            End If

        Else

            LoadFileID3Ex = False

        End If
    End If

End Function
'#End Region

'#Region "Callback"
Public Function SetCallbackFunc(ByVal Messages As TCallbackMessage, ByVal UserData _
    As Long) As Boolean

    If oPlayer <> 0 Then

        SetCallbackFunc = CBool(zplay_SetCallbackFunc(oPlayer, AddressOf _
            TCallbackFunc, Messages, UserData))

    End If

End Function
'#End Region

'#Region "Beat-Per-Minute"
Public Function DetectBPM(ByVal Method As TBPMDetectionMethod) As Long

    If oPlayer <> 0 Then

        DetectBPM = zplay_DetectBPM(oPlayer, Method)

    End If

End Function

Public Function DetectFileBPM(ByVal FileName As String, ByVal Format As _
    TStreamFormat, ByVal Method As TBPMDetectionMethod) As Long

    Dim pStr As Long

    If oPlayer <> 0 Then

        pStr = StrPtr(FileName)
        DetectFileBPM = zplay_DetectFileBPMW(oPlayer, pStr, Format, Method)

    End If

End Function
'#End Region

'#Region "FFT Graph and FFT values"
Public Function GetFFTData(ByVal FFTPoints As Long, ByVal FFTWindow As TFFTWindow, _
    ByRef HarmonicNumber As Long, ByRef HarmonicFreq() As Long, ByRef _
    LeftAmplitude() As Long, ByRef RightAmplitude() As Long, ByRef LeftPhase() As _
    Long, ByRef RightPhase() As Long) As Boolean

    If oPlayer <> 0 Then

        GetFFTData = CBool(zplay_GetFFTData(oPlayer, FFTPoints, FFTWindow, _
            HarmonicNumber, HarmonicFreq, LeftAmplitude, RightAmplitude, _
            LeftPhase, RightPhase))

    End If

End Function

Public Function DrawFFTGraphOnHDC(ByVal hdc As Long, ByVal X As Long, ByVal Y As _
    Long, ByVal Width As Long, ByVal Height As Long) As Boolean

    If oPlayer <> 0 Then

        DrawFFTGraphOnHDC = CBool(zplay_DrawFFTGraphOnHDC(oPlayer, hdc, X, Y, _
            Width, Height))

    End If

End Function

Public Function DrawFFTGraphOnHWND(ByVal hwnd As Long, ByVal X As Long, ByVal Y As _
    Long, ByVal Width As Long, ByVal Height As Long) As Boolean

    If oPlayer <> 0 Then

        DrawFFTGraphOnHWND = CBool(zplay_DrawFFTGraphOnHWND(oPlayer, hwnd, X, Y, _
            Width, Height))

    End If

End Function

Public Function SetFFTGraphParam(ByVal ParamID As TFFTGraphParamID, ByVal _
    ParamValue As Long) As Boolean

    If oPlayer <> 0 Then

        SetFFTGraphParam = CBool(zplay_SetFFTGraphParam(oPlayer, ParamID, _
            ParamValue))

    End If

End Function

Public Function GetFFTGraphParam(ByVal ParamID As TFFTGraphParamID) As Long

    If oPlayer <> 0 Then

        GetFFTGraphParam = zplay_GetFFTGraphParam(oPlayer, ParamID)

    End If

End Function
'#End Region

'#Region "Center and side cut"
Public Function StereoCut(ByVal Enable As Boolean, ByVal OutputCenter As Boolean, _
    ByVal BassToSides As Boolean) As Boolean

    Dim fOutputCenter As Long
    Dim fBassToSides As Long
    Dim fEnable As Long

    If oPlayer <> 0 Then

        fOutputCenter = 0
        fBassToSides = 0
        fEnable = 0

        If OutputCenter Then

            fOutputCenter = 1

        End If

        If BassToSides Then

            fBassToSides = 1

        End If

        If Enable Then

            fEnable = 1

        End If

        StereoCut = CBool(zplay_StereoCut(oPlayer, fEnable, fOutputCenter, _
            fBassToSides))

    End If

End Function
'#End Region

'#Region "Channel mixing"
Public Function MixChannels(ByVal Enable As Boolean, ByVal LeftPercent As Long, _
    ByVal RightPercent As Long) As Boolean

    If oPlayer <> 0 Then
        If Enable Then

            MixChannels = CBool(zplay_MixChannels(oPlayer, 1, LeftPercent, _
                RightPercent))

        Else

            MixChannels = CBool(zplay_MixChannels(oPlayer, 0, LeftPercent, _
                RightPercent))

        End If
    End If

End Function
'#End Region

'#Region "VU Data"
Public Sub GetVUData(ByRef LeftChannel As Long, ByRef RightChannel As Long)

    If oPlayer <> 0 Then

        Call zplay_GetVUData(oPlayer, LeftChannel, RightChannel)

    End If

End Sub
'#End Region

'#Region "Status and Info"
Public Sub GetStreamInfo(ByRef Info As TStreamInfo)

    Dim tmp As TStreamInfo_Internal

    If oPlayer <> 0 Then

        Call zplay_GetStreamInfoW(oPlayer, tmp)

        Info.BitRate = tmp.BitRate
        Info.ChannelNumber = tmp.ChannelNumber
        Info.SamplingRate = tmp.SamplingRate
        Info.VBR = tmp.VBR
        Info.Length = tmp.Length
        Info.Description = PointerToString(tmp.Description)

    End If

End Sub

Public Sub GetStatus(ByRef status As TStreamStatus)

    If oPlayer <> 0 Then

        Call zplay_GetStatus(oPlayer, status)

    End If

End Sub

Public Sub GetDynamicStreamLoad(ByRef StreamLoadInfo As TStreamLoadInfo)

    If oPlayer <> 0 Then

        Call zplay_GetDynamicStreamLoad(oPlayer, StreamLoadInfo)

    End If

End Sub
'#End Region

'#Region "Wave Out and Wave In Info"
Public Function EnumerateWaveOut() As Long

    If oPlayer <> 0 Then

        EnumerateWaveOut = zplay_EnumerateWaveOut(oPlayer)

    End If

End Function

Public Function GetWaveOutInfo(ByVal Index As Long, ByRef Info As TWaveOutInfo) As _
    Boolean

    Dim tmp As TWaveOutInfo_Internal

    If oPlayer <> 0 Then
        If zplay_GetWaveOutInfoW(oPlayer, Index, tmp) = 0 Then

            GetWaveOutInfo = False

        Else

            Info.Channels = tmp.Channels
            Info.DriverVersion = tmp.DriverVersion
            Info.Formats = tmp.Formats
            Info.ManufacturerID = tmp.ManufacturerID
            Info.ProductID = tmp.ProductID
            Info.Support = tmp.Support
            Info.ProductName = PointerToString(tmp.ProductName)
            GetWaveOutInfo = True

        End If
    End If

End Function

Public Function SetWaveOutDevice(ByVal Index As Long) As Boolean

    If oPlayer <> 0 Then

        SetWaveOutDevice = CBool(zplay_SetWaveOutDevice(oPlayer, Index))

    End If

End Function

Public Function EnumerateWaveIn() As Long

    If oPlayer <> 0 Then

        EnumerateWaveIn = zplay_EnumerateWaveIn(oPlayer)

    End If

End Function

Public Function GetWaveInInfo(ByVal Index As Long, ByRef Info As TWaveInInfo) As _
    Boolean

    Dim tmp As TWaveInInfo_Internal

    If oPlayer <> 0 Then
        If zplay_GetWaveInInfoW(oPlayer, Index, tmp) = 0 Then

            GetWaveInInfo = False

        Else

            Info.Channels = tmp.Channels
            Info.DriverVersion = tmp.DriverVersion
            Info.Formats = tmp.Formats
            Info.ManufacturerID = tmp.ManufacturerID
            Info.ProductID = tmp.ProductID
            Info.ProductName = PointerToString(tmp.ProductName)
            GetWaveInInfo = True

        End If
    End If

End Function
' #End Region

'#Region "Settings"
Public Function SetSettings(ByVal SettingID As TSettingID, ByVal Value As Long) As _
    Long

    If oPlayer <> 0 Then

        SetSettings = zplay_SetSettings(oPlayer, SettingID, Value)

    End If

End Function

Public Function GetSettings(ByVal SettingID As TSettingID) As Long

    If oPlayer <> 0 Then

        GetSettings = zplay_GetSettings(oPlayer, SettingID)

    End If

End Function
'#End Region

' ----==== Hilfsfunktion ====----
Private Function StartUpGDIPlus() As Long

    Dim GdipStartupInput As GDIPlusStartupInput
    Dim GdipStartupOutput As GdiplusStartupOutput

    GdipStartupInput.GdiPlusVersion = 1
    StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, GdipStartupOutput)

End Function

Private Function ShutdownGDIPlus() As Long

    If GdiplusShutdown(GdipToken) = 0 Then

        GdipToken = 0
        ShutdownGDIPlus = 0

    End If

End Function

Private Function StreamToPicture(ByVal PicStream As IUnknown) As StdPicture

    Dim lBitmap As Long
    Dim hBitmap As Long

    If StartUpGDIPlus = 0 Then
        If GdipLoadImageFromStream(PicStream, lBitmap) = 0 Then
            If GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0) = 0 Then

                Set StreamToPicture = HandleToPicture(hBitmap, vbPicTypeBitmap)

            End If

            Call GdipDisposeImage(lBitmap)

        End If

        Call ShutdownGDIPlus

    End If

End Function

Private Function HandleToPicture(ByVal hGDIHandle As Long, ByVal ObjectType As _
    PictureTypeConstants, Optional ByVal hPal As Long = 0) As StdPicture

    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As IID
    Dim oPicture As IPicture

    With tPictDesc

        .cbSizeOfStruct = Len(tPictDesc)
        .picType = ObjectType
        .hgdiObj = hGDIHandle
        .hPalOrXYExt = hPal

    End With

    With IID_IPicture

        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB

    End With

    Call OleCreatePictureIndirect(tPictDesc, IID_IPicture, True, oPicture)

    Set HandleToPicture = oPicture

End Function

Private Function PointerToString(ByVal lngPtr As Long) As String

    Dim lngLen As Long
    Dim strTemp As String

    If lngPtr Then

        lngLen = lstrlenW(lngPtr) * 2

        If lngLen Then

            strTemp = Space(lngLen)
            CopyMemory ByVal strTemp, ByVal lngPtr, lngLen
            PointerToString = Replace(strTemp, Chr$(0), vbNullString)

        End If
    End If

End Function
