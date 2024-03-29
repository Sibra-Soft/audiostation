VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mp3Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Class:            MP3Info
' Description:      Gets header and ID3Tag details from a MP3 file
'
' Date Changed:     05-10-2021
' Date Created:     04-10-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Private Const FILE_BEGIN = 0
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3

Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private m_MPEGVer As Single
Private m_VerStr As String
Private m_Layers As Long
Private m_Protected As Boolean
Private m_BitRate As Long
Private m_Frequency As Long
Private m_ChannelMode As String
Private m_CopyRighted As Boolean
Private m_Original As Boolean
Private m_Emphasis As String
Private m_Padding As Byte
Private m_ValidMP3 As Boolean
Private m_FileName As String
Private m_FileSize As Long
Private m_SongLength As Long
Private m_NumFrames As Long
Private m_FrameLength As Long
Private m_VBR As Boolean
Private m_HeaderOffset As Long
Private m_ActSize As Long

Public Artist As String
Public Title As String
Public Album As String
Public Year As String
Public Genre As String

Private Type ID3v1Tag
  id As String * 3
  Title As String * 30
  Artist As String * 30
  Album As String * 30
  Year As String * 4
  comment As String * 30
  Genre As Byte
End Type

Private Version As Byte

Private Freq0 As Variant, Freq2 As Variant, Freq3 As Variant
Private BR3 As Variant, BR2 As Variant, BR1 As Variant, BR31 As Variant, BR32 As Variant

Private i As Integer
Private strEmptyString As String
Private B() As Byte
Private B2 As Byte

Private s As String
Private Function GetID3v1Tags()
On Error GoTo ErrorHandler
Dim strFilename As String
Dim lngFilesize As Long

Dim FN As Integer
Dim lngHeaderPosition As Long
Dim Tag1 As ID3v1Tag
Dim Tag2 As String

FN = FreeFile

Artist = vbNullString
Title = vbNullString
Album = vbNullString
Year = vbNullString
Genre = vbNullString

Open m_FileName For Binary As #FN
    lngFilesize = LOF(FN)
    
    Get #FN, 1, B2
        
    If B2 <> 255 Then '(255 is where an ID3v2 header should start)
        If B2 <> 73 Then
            'Exit Function
        End If
    End If
     
    lngHeaderPosition = 1
    Get #FN, 2, B2
    If (B2 < 250 Or B2 > 251) Then
        'We have an ID3v2 tag
        If B2 = 68 Then
            Get #FN, 3, B2
            If B2 = 51 Then
                Dim r As Double
                Get #FN, 4, Version
                Get #FN, 7, B2
                r = B2 * 20917152
                Get #FN, 8, B2
                r = r + (B2 * 16384)
                Get #FN, 9, B2
                r = r + (B2 * 128)
                Get #FN, 10, B2
                r = r + B2
                If r > lngFilesize Or r > 2147483647 Then
                    Exit Function
                End If
                Tag2 = Space$(r)
                Get #FN, 11, Tag2
                lngHeaderPosition = r + 11
            End If
        End If
    Else
        'ID3v2 tag is missing
    End If
    
    'ID3v1 tag
    Get #FN, lngFilesize - 127, Tag1.id
    
    If Tag1.id = "TAG" Then 'If "TAG" is present, then we have a valid ID3v1 tag and will extract all available ID3v1 info from the file
        Get #FN, , Tag1.Title   'Always limited to 30 characters
        Get #FN, , Tag1.Artist  'Always limited to 30 characters
        Get #FN, , Tag1.Album   'Always limited to 30 characters
        Get #FN, , Tag1.Year    'Always limited to 4 characters
        Get #FN, , Tag1.comment 'Always limited to 30 characters
        Get #FN, , Tag1.Genre   'Always limited to 1 byte (?)
    
        'Trim$(Tag1.title)
        'Trim$(Tag1.artist)
        'Trim$(Tag1.album)
        'Trim$(Tag1.year)
        'Trim$(Tag1.comment)
        'Tag1.genre
    End If
    
    If Tag2 <> strEmptyString Then GetID3v2Tags (Tag2) 'Pass the Id3v2 TagId to the GetID3v2Tag1 function
Close
Exit Function
        
ErrorHandler:
    'MsgBox "Error reading file"
    err.Clear
    Close
    Resume Next
End Function
Private Function GetID3v2Tags(Tag2 As String) As Boolean

On Error GoTo ErrorHandler

Dim TitleField As String
Dim ArtistField As String
Dim AlbumField As String
Dim YearField As String
Dim GenreField As String
Dim FieldSize As Long
Dim SizeOffset As Long
Dim FieldOffset As Long
Dim TrackNbr As String
Dim SituationField As String

Select Case Version
    Case 2 'ID3v2.2
        'Set the fieldnames for version 2.0
        TitleField = "TT2"
        ArtistField = "TOA"
        AlbumField = "TAL"
        YearField = "TYE"
        GenreField = "TCO"
        FieldOffset = 7
        SizeOffset = 5
        TrackNbr = "TRCK"
   
    Case 3 'ID3v2.3
        'Set the fieldnames for version 3.0
        TitleField = "TIT2"
        ArtistField = "TPE1"
        AlbumField = "TALB"
        YearField = "TYER"
        GenreField = "TCON"
        TrackNbr = "TRCK"
   
        FieldOffset = 11
        SizeOffset = 7
    Case Else
        'We don't have a valid ID3v2 tag, so bail
        Exit Function
End Select

' Title
i = InStr(Tag2, TitleField)
If i > 0 Then
   'read the title
   FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
   If Version = 3 Then
      'check for compressed or encrypted field
      B2 = Asc(Mid$(Tag2, i + 9))
      If (B2 And 128) = True Or (B2 And 64) = True Then GoTo ReadAlbum
   End If
   
   Title = Mid$(Tag2, i + FieldOffset, FieldSize)
End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract album title
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadAlbum:
 i = InStr(Tag2, AlbumField)
 If i > 0 Then
    FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
    If Version = 3 Then
       'check for compressed or encrypted field
       B2 = Asc(Mid$(Tag2, i + 9))
       If (B2 And 128) = 128 Or (B2 And 64) = 64 Then GoTo ReadArtist
    End If
    
    Album = Mid$(Tag2, i + FieldOffset, FieldSize)
 End If
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract artist name
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadArtist:
i = InStr(Tag2, ArtistField)
If i > 0 Then
   FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
   If Version = 3 Then
      'check for compressed or encrypted field
      B2 = Asc(Mid$(Tag2, i + 9))
      If (B2 And 128) = 128 Or (B2 And 64) = 64 Then GoTo ReadYear
   End If
    Artist = Mid$(Tag2, i + FieldOffset, FieldSize)
End If

ReadYear:
i = InStr(Tag2, YearField)
If i > 0 Then
   FieldSize = Asc(Mid$(Tag2, i + SizeOffset)) - 1
   If Version = 3 Then
      'check for compressed or encrypted field
      B2 = Asc(Mid$(Tag2, i + 9))
      If (B2 And 128) = 128 Or (B2 And 64) = 64 Then GoTo ReadGenre
   End If
   
   Year = Mid$(Tag2, i + FieldOffset, FieldSize)
End If

ReadGenre:
i = InStr(Tag2, GenreField)

ReadTrackNbr:
i = InStr(Tag2, TrackNbr)

Done:
    Exit Function

ErrorHandler:
    err.Clear
    Resume Next
End Function
Private Sub Class_Initialize()
 Freq0 = Array(11025, 12000, 8000, 0)
 Freq2 = Array(22050, 24000, 26000, 0)
 Freq3 = Array(44100, 48000, 32000, 0)
 BR3 = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448, 0)
 BR2 = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384, 0)
 BR1 = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 0)
 BR31 = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256, 0)
 BR32 = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160, 0)
End Sub

Private Sub Class_Terminate()
'Erase m_GenreArray
End Sub
Public Function FileExists(ByVal sFile As String) As Boolean
 Dim eAttr As Long
 On Error Resume Next
 eAttr = GetAttr(sFile)
 FileExists = (err.Number = 0) And ((eAttr And vbDirectory) = 0)
 On Error GoTo 0
End Function
'Public interface to member varaibles
Public Property Get Filename() As String
 Filename = m_FileName
End Property
Public Property Let Filename(ByVal NewVal As String)
    If FileExists(NewVal) Then
        m_FileName = NewVal
        ReadMp3Binary
    Else
        err.Raise 53
    End If
End Property

Public Property Get ValidMP3() As Boolean
 ValidMP3 = m_ValidMP3
End Property
Public Property Get Padding() As Byte
 Padding = m_Padding
End Property
Public Property Get FrameLength() As Long
 FrameLength = m_FrameLength
End Property
Public Property Get FileSize() As Long
 FileSize = m_FileSize
End Property

Public Property Get BitRate() As Long
 BitRate = m_BitRate
End Property

Public Property Get Frequency() As Long
 Frequency = m_Frequency
End Property

Public Property Get SongLength() As Long
 SongLength = m_SongLength
End Property

Public Property Get FormattedTime() As String 'Format track time into HH:MM:SS (Or MM:SS)
 FormattedTime = Format$(CDate(m_SongLength / 86400), "nn:ss")
End Property

Public Property Get NumFrames() As Long
 NumFrames = m_NumFrames
End Property

Public Property Get isVBR() As Boolean
 isVBR = m_VBR
End Property

Public Property Get MPEGVer() As Single
 MPEGVer = m_MPEGVer
End Property

Public Property Get Layers() As Byte
 Layers = m_Layers
End Property
Public Property Get MPEGLayer() As String
 MPEGLayer = "MPEG " & m_VerStr & " Layer " & m_Layers
End Property
Public Property Get Protected() As Boolean
 Protected = m_Protected
End Property

Public Property Get ChannelMode() As String
 ChannelMode = m_ChannelMode
End Property

Public Property Get CopyRighted() As Boolean
 CopyRighted = m_CopyRighted
End Property

Public Property Get Original() As Boolean
 Original = m_Original
End Property

Public Property Get HeaderOffset() As Long
 HeaderOffset = m_HeaderOffset
End Property

Public Property Get Emphasis() As String
 Emphasis = m_Emphasis
End Property
Private Function StringToSize(ByVal str As String, Optional ByVal IsHeaderSize As Boolean = True) As Double
 Dim i As Long
 Dim num As Double
 Dim Base As Double
 num = 0
 Base = 256 * (2 ^ CDbl(IsHeaderSize))
 For i = Len(str) To 1 Step -1
  num = num + CDbl(Asc(Mid$(str, i, 1))) * Base ^ CDbl(Len(str) - i)
 Next
 StringToSize = num
End Function

Public Sub ReadMp3()
 Dim FNum As Integer
 Dim ReadPos As Long
 Dim GetHeader(3) As Byte
 Dim VBRHeader(11) As Byte
 Dim HV0 As Byte
 Dim i As Long, a(3) As Byte
 Dim IDTest As String
 Dim TagSize As Long
 IDTest = Space$(10)                        ' Give 10 blank bytes to the variable

 Call ClearInfo
 On Error GoTo BadFile
 FNum = FreeFile
 Open m_FileName For Binary As FNum
 m_FileSize = LOF(FNum)
 Get #FNum, 1, IDTest
 If Left$(IDTest, 3) = "ID3" And _
   Asc(Mid$(IDTest, 4, 1)) < &HFF& And _
   Asc(Mid$(IDTest, 5, 1)) < &HFF& And _
   Asc(Mid$(IDTest, 7, 1)) < &H80& And _
   Asc(Mid$(IDTest, 8, 1)) < &H80& And _
   Asc(Mid$(IDTest, 9, 1)) < &H80& And _
   Asc(Mid$(IDTest, 10, 1)) < &H80& Then    ' ID3 header found
  TagSize = StringToSize(Right$(IDTest, 4))
  ReadPos = TagSize - 1
 End If
 Do                                         'Scan until a valid header is found
  Do                                        'Scan through the file looking for possible headers
   ReadPos = ReadPos + 1
   Seek #FNum, ReadPos
   Get #FNum, , GetHeader()
  Loop Until IsHeader(GetHeader())
 Loop Until ValidHeader(GetHeader(), ReadPos) Or (ReadPos = (m_FileSize - 5))
 'Read in variable bitrate header
 If ReadPos < (m_FileSize - 16) Then
  Seek #FNum, ReadPos + 36
  Get #FNum, , VBRHeader
 End If
 Close FNum
 HV0 = (GetHeader(1) \ &H8) And &H3         'MPEG version (B)
 m_VBR = (ChrW$(VBRHeader(0)) & ChrW$(VBRHeader(1)) & ChrW$(VBRHeader(2)) & ChrW$(VBRHeader(3))) = "Xing"
 If m_VBR Then                              'Variable BitRate encoded MP3 ( !!UNTESTED!! )
  If (VBRHeader(7) And &H1) = &H1 Then      'FRAMES_FLAG
   For i = 3 To 0 Step -1
    a(3 - i) = VBRHeader(i + 8)
   Next
   CopyMemory m_NumFrames, a(0), 4
   m_NumFrames = m_NumFrames - 1
   m_BitRate = ((m_FileSize / m_NumFrames) * m_Frequency) / (1000 * IIf(HV0 = &H3, 12, 144))
   m_BitRate = m_BitRate \ 12
   m_SongLength = ((8 * m_ActSize) \ 1000) \ m_BitRate
   m_ValidMP3 = True
  Else
   m_ValidMP3 = False
  End If
 End If
BadFile:
 If Not m_ValidMP3 Then Call ClearInfo
End Sub
Private Function IsHeader(inArr() As Byte) As Boolean 'Tests for MPEG header sync information
 IsHeader = (inArr(0) = &HFF) And (((inArr(1) \ &H20) And &H7) = &H7)
End Function
Private Function ValidHeader(inArr() As Byte, ReadPos As Long) As Boolean
'if exit function occurs, this is not
'  a valid header

 Dim HeadVals(11) As Byte

 ' 44100 / 1130.76923076923=39 frames per sec

 '32-bit MP3 frame header construction:
 '87654321 87654321 87654321 87654321
 'AAAAAAAA AAABBCCD EEEEFFGH IIJJKLMM

 'A = Sync information: 11 bits all set to 1,
 'used to find header position
 'BB=MPEG Ver, CC=Layer, D=Protection
 'EEEE=Bitrate, FF=Frequency, G=Padding Bit
 'H=Private Bit, II=Channel Mode
 'JJ=Mode extension (used only with joint stereo, to conjoin channel data)
 'K=Copyright Bit, L= Original Bit
 'MM=Emphasis (obsolete?)
 HeadVals(0) = (inArr(1) \ &H8) And &H3     'MPEG version (B)
 HeadVals(1) = (inArr(1) \ &H2) And &H3     'Layer description (C)
 HeadVals(2) = inArr(1) And &H1             'Protection bit (D)
 HeadVals(3) = (inArr(2) \ &H10) And &HF    'Bitrate index (E)
 HeadVals(4) = (inArr(2) \ &H4) And &H3     'Sample rate (F)
 HeadVals(5) = (inArr(2) \ &H2) And &H1     'Padding bit (G)
 HeadVals(6) = inArr(2) And &H1             'Private bit (H)
 HeadVals(7) = (inArr(3) \ &H40) And &H3    'Channel mode (I)
 HeadVals(8) = (inArr(3) \ &H10) And &H3    'Mode extension (J)
 HeadVals(9) = (inArr(3) \ &H8) And &H1     'Copyright bit (K)
 HeadVals(10) = (inArr(3) \ &H4) And &H1    'Original bit (L)
 HeadVals(11) = inArr(3) And &H3            'Emphasis (M)

 Select Case HeadVals(0)                    'Select frequency based on MPEG version
  Case 0                                    'MPEG 2.5
   m_Frequency = Freq0(HeadVals(4))
  Case 1: m_Frequency = 0                   'Reserved
  Case 2
   m_Frequency = Freq2(HeadVals(4))
  Case 3
   m_Frequency = Freq3(HeadVals(4))
 End Select
 If m_Frequency = 0 Then Exit Function

 If HeadVals(0) = &H3 Then                  'MPEG 1.0
  Select Case HeadVals(1)
   Case &H3
    m_BitRate = BR3(HeadVals(3))
   Case &H2
    m_BitRate = BR2(HeadVals(3))
   Case &H1
    m_BitRate = BR1(HeadVals(3))
  End Select
 Else                                       'MPEG 2.0, 2.5 or undefined
  Select Case HeadVals(1)
   Case &H3
    m_BitRate = BR31(HeadVals(3))
   Case &H2, &H1
    m_BitRate = BR32(HeadVals(3))
  End Select
 End If
 If m_BitRate = 0 Then Exit Function
 m_ActSize = m_FileSize - ReadPos           '- 'TagLen
 m_NumFrames = (m_ActSize) \ ((m_BitRate * 144000) \ (m_Frequency))

 m_HeaderOffset = ReadPos
 m_SongLength = ((8 * m_ActSize) \ 1000) \ m_BitRate
 Select Case HeadVals(0) + 1
  Case 1: m_MPEGVer = 2.5: m_VerStr = "2.5"
  Case 2: m_MPEGVer = 0: m_VerStr = "0.0"
  Case 3: m_MPEGVer = 2: m_VerStr = "2.0"
  Case 4: m_MPEGVer = 1: m_VerStr = "1.0"
 End Select
 Select Case HeadVals(1) + 1
  Case 1: m_Layers = 0
  Case 2: m_Layers = 3
  Case 3: m_Layers = 2
  Case 4: m_Layers = 1
 End Select
 If m_MPEGVer = 0 Or m_Layers = 0 Then Exit Function
 m_Padding = HeadVals(5)
 If m_Layers = 1 Then
  m_FrameLength = (12 * (m_BitRate * 1000&) \ m_Frequency + m_Padding) * 4
 Else
  m_FrameLength = (144 * (m_BitRate * 1000&) \ m_Frequency) + m_Padding
 End If

 Select Case HeadVals(7) + 1
  Case 1: m_ChannelMode = "Stereo"
  Case 2: m_ChannelMode = "Joint Stereo"
  Case 3: m_ChannelMode = "Dual Channel"
  Case 4: m_ChannelMode = "Single Channel"
 End Select
 m_Protected = HeadVals(2) = &H0
 m_CopyRighted = HeadVals(9) = &H1
 m_Original = HeadVals(10) = &H1
 Select Case HeadVals(11) + 1
  Case 1: m_Emphasis = "None"
  Case 2: m_Emphasis = "50/15 ms"
  Case 3: m_Emphasis = "Reserved"
  Case 4: m_Emphasis = "CCIT J.17"
 End Select
 m_ValidMP3 = True
 ValidHeader = True
End Function

Private Sub ClearInfo() 'Clears out all current information
 m_MPEGVer = 0
 m_Layers = 0
 m_Protected = False
 m_BitRate = 0
 m_Frequency = 0
 m_ChannelMode = ""
 m_CopyRighted = False
 m_Original = False
 m_ValidMP3 = False
 m_FileSize = 0
 m_SongLength = 0
 m_NumFrames = 0
 m_VBR = False
 m_HeaderOffset = 0
End Sub
Private Function ReadFileBinary(ByVal sFilePath As String) As Byte()
 Dim hFile As Long
 hFile = FreeFile
 Open sFilePath For Binary As #hFile
 m_FileSize = LOF(hFile)
 ReDim ReadFileBinary(1 To m_FileSize)
 Get #hFile, , ReadFileBinary
 Close #hFile
End Function
Public Function StringToSizeBin(Arr() As Byte, Optional ByVal IsHeaderSize As Boolean = True) As Double
 Dim i As Long
 Dim num As Long                            ' Double
 Dim Base As Long                           'Double
 Base = 256 * (2 ^ CDbl(IsHeaderSize))
 ' Debug.Print "Ish", IsHeaderSize, "Base", Base
 For i = 4 To 1 Step -1
  num = num + Arr(i) * Base ^ (4 - i)
 Next
 StringToSizeBin = num
End Function

Private Sub GetReadPos(RP As Long)
 Dim a(1 To 4) As Byte
 If Chr$(B(1)) = "I" And Chr$(B(2)) = "D" And Chr$(B(3)) = "3" Then
  CopyMemory a(1), B(7), 4
  RP = StringToSizeBin(a) - 1
  '  Debug.Print Hex(RP)
 End If
End Sub

Public Sub ReadMp3Binary()
 Dim ReadPos As Long
 Dim GetHeader(3) As Byte
 Dim VBRHeader(11) As Byte
 Dim HV0 As Byte
 Dim i As Long, a(3) As Byte
 ' Dim TagSize As Long
 Call ClearInfo
 B = ReadFileBinary(m_FileName)
 GetReadPos ReadPos
 Do                                         'Scan until a valid header is found
  Do                                        'Scan through the file looking for possible headers
   ReadPos = ReadPos + 1
   CopyMemory GetHeader(0), B(ReadPos), 4
  Loop Until IsHeader(GetHeader())
 Loop Until ValidHeader(GetHeader(), ReadPos) Or (ReadPos = (m_FileSize - 5))
 If ReadPos < (m_FileSize - 16) Then
  CopyMemory VBRHeader(0), B(ReadPos + 36), 12
 End If
 HV0 = (GetHeader(1) \ &H8) And &H3         'MPEG version (B)
 m_VBR = (ChrW$(VBRHeader(0)) & ChrW$(VBRHeader(1)) & ChrW$(VBRHeader(2)) & ChrW$(VBRHeader(3))) = "Xing"
 If m_VBR Then                              'Variable BitRate encoded MP3 ( !!UNTESTED!! )
  If (VBRHeader(7) And &H1) = &H1 Then      'FRAMES_FLAG
   For i = 3 To 0 Step -1
    a(3 - i) = VBRHeader(i + 8)
   Next
   CopyMemory m_NumFrames, a(0), 4
   m_NumFrames = m_NumFrames - 1
   m_BitRate = (((m_FileSize / m_NumFrames) * m_Frequency) / (1000 * IIf(HV0 = &H3, 12, 144))) \ 12
   m_SongLength = ((8 * m_ActSize) \ 1000) \ m_BitRate
   m_ValidMP3 = True
   
   Call GetID3v1Tags
  Else
   m_ValidMP3 = False
  End If
 End If
 Exit Sub
BadFile:
 If Not m_ValidMP3 Then Call ClearInfo
End Sub
Public Sub RemovePreHeader()
 Dim hFile As Long
 If m_HeaderOffset > 1 Then
  CopyMemory B(1), B(m_HeaderOffset), m_FileSize - m_HeaderOffset + 1
  ReDim Preserve B(1 To m_FileSize - m_HeaderOffset + 1)
  Kill m_FileName
  hFile = FreeFile
  Open m_FileName For Binary As hFile
  Put #hFile, , B
  Close hFile
  ReadMp3Binary
 End If
End Sub

Public Sub RemoveV1Tag()
 Dim T As String
 Dim i As Long
 Dim hFile As Long
 For i = 127 To 125 Step -1
  T = T & Chr$(B(m_FileSize - i))
 Next
 If T <> "TAG" Then Exit Sub
 hFile = CreateFile(m_FileName, GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
 Call SetFilePointer(hFile, m_FileSize - 128, 0&, FILE_BEGIN)
 Call SetEndOfFile(hFile)
 Call CloseHandle(hFile)
 ReadMp3Binary
End Sub

