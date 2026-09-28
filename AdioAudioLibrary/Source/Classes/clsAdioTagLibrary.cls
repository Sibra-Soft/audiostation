VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdioTagLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type ID3Tag
    Artist As String
    Album As String
    Title As String
    SongYear As String
    Genre As String
    Comment As String
    TrackNr As String
    EncodedBy As String
    LinkTo As String
    Copyright As String
    Composer As String
    OrigArtist As String
End Type

Public Type MPEGInfo
    MPEGVersion As String
    HasCRC As Boolean
    Bitrate As Long     'kbit / sec (where k = 1000, not 1024)
    Frequency As Long   'hz
    HasVBR As Boolean
    ChannelMode As String
    IsCopyrighted As Boolean
    IsOriginal As Boolean
    HasEmphasis As Boolean
    length As Long      'sec
    ID3v1Version As Long    '-1 = no tag
    ID3v2Version As Long    '-1 = no tag
    FileSize As Long    'bytes
End Type

Public Enum ID3v2_Version
    VERSION_2_3
    VERSION_2_4
End Enum

Public ID3_GENRE_COUNT As Long

Private Type v1Tag
    Identifier(2) As Byte
    Title(29) As Byte
    Artist(29) As Byte
    Album(29) As Byte
    SongYear(3) As Byte
    Comment(29) As Byte
    Genre As Byte
End Type

'***** ID3v2 Types *****

'Tag header for all versions (2.2 - 2.4)
'The v 2.4 footer also uses this structure, since it's a copy of the header.
Private Type v2TagHeader
    Identifier(2) As Byte
    Version(1) As Byte
    flags As Byte
    Size(3) As Byte
End Type

'Extended tag header for v 2.3
Private Type v2_3ExtHeader
    Size(3) As Byte
    ExtendedFlags(1) As Byte
    PaddingSize(3) As Byte
End Type

'Extended tag header for v 2.4
Private Type v2_4ExtHeader
    Size(3) As Byte
    NumFlags As Byte
    ExtendedFlags As Byte
End Type

'Frame header for v 2.2
Private Type v2_2FrameHeader
    Ident(2) As Byte
    Size(2) As Byte
End Type

'Frame header for v 2.3 and v 2.4
Private Type v2_34FrameHeader
    Ident(3) As Byte
    Size(3) As Byte
    flags(1) As Byte
End Type

'A v 2.x frame in memory.
Private Type v2MemFrame
    Id As String
    Size As Long
    Data() As Byte
End Type

Private Type v2TagHeaderFlags
    bUnsynch As Boolean
    bExtHeader As Boolean
    bFooter As Boolean
    bUpdateTag As Boolean
End Type

Private Type v2TagSizes
    ReadSize As Long
    ExtHeaderSize As Long
    FooterSize As Long
End Type

Private Enum v2_StrEncoding
    ENC_ISO = 0
    ENC_UNICODE_UTF16_BOM = 1
    ENC_UNICODE_UTF16 = 2
    ENC_UNICODE_UTF8 = 3
End Enum

'The frames stored in memory, used when updating a currently existing v2 tag.
Private mMemFrames() As v2MemFrame, mMemFraCnt As Long

Private Const cGenreString As String = _
"Blues#Classic Rock#Country#Dance#Disco#Funk#Grunge#Hip-Hop#Jazz#Metal#New Age#Oldies#Other#Pop#" & _
"R&B#Rap#Reggae#Rock#Techno#Industrial#Alternative#Ska#Death Metal#Pranks#Soundtrack#Euro-Techno#" & _
"Ambient#Trip-Hop#Vocal#Jazz+Funk#Fusion#Trance#Classical#Instrumental#Acid#House#Game#Sound Clip#" & _
"Gospel#Noise#Alt. Rock#Bass#Soul#Punk#Space#Meditative#Instrumental Pop#Instrumental Rock#Ethnic#" & _
"Gothic#Darkwave#Techno-Industrial#Electronic#Pop-Folk#Eurodance#Dream#Southern Rock#Comedy#Cult#" & _
"Gangsta#Top 40#Christian Rap#Pop/Funk#Jungle#Native American#Cabaret#New Wave#Psychadelic#Rave#" & _
"Showtunes#Trailer#Lo-Fi#Tribal#Acid Punk#Acid Jazz#Polka#Retro#Musical#Rock & Roll#Hard Rock#Folk#" & _
"Folk-Rock#National Folk#Swing#Fast Fusion#Bebob#Latin#Revival#Celtic#Bluegrass#Avantgarde#Gothic Rock#" & _
"Progressive Rock#Psychedelic Rock#Symphonic Rock#Slow Rock#Big Band#Chorus#Easy Listening#Acoustic#" & _
"Humour#Speech#Chanson#Opera#Chamber Music#Sonata#Symphony#Booty Bass#Primus#Porn Groove#Satire#Slow Jam#" & _
"Club#Tango#Samba#Folklore#Ballad#Power Ballad#Rhythmic Soul#Freestyle#Duet#Punk Rock#Drum Solo#A capella#" & _
"Euro-House#Dance Hall#Goa#Drum & Bass#Club-House#Hardcore#Terror#Indie#BritPop#Negerpunk#Polsk Punk#" & _
"Beat#Christian Gangsta#Heavy Metal#Black Metal#Crossover#Contemporary C#Christian Rock#Merengue#Salsa#" & _
"Thrash Metal#Anime#JPop#SynthPop"
Private GenreField() As String, bBuiltGenre As Boolean

'API
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = &H3

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const INVALID_HANDLE_VALUE = -1

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Function ReadID3v1(ByVal FileName As String, ByRef outTag As ID3Tag) As Boolean
Dim fh As Long, fsize As Long
Dim RdTag As v1Tag

fh = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
If fh = INVALID_HANDLE_VALUE Then Exit Function
fsize = GetFileSize(fh, 0)
If fsize < 128 Then
    CloseHandle fh
    Exit Function
End If

SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
ReadFile fh, RdTag, Len(RdTag), 0, ByVal 0
CloseHandle fh

If Data2String(VarPtr(RdTag.Identifier(0)), 3, ENC_ISO) = "TAG" Then
    ResetTag outTag

    'An ID3v1 tag is present.
    outTag.Artist = Trim$(Data2String(VarPtr(RdTag.Artist(0)), 30, ENC_ISO))
    outTag.Album = Trim$(Data2String(VarPtr(RdTag.Album(0)), 30, ENC_ISO))
    outTag.Title = Trim$(Data2String(VarPtr(RdTag.Title(0)), 30, ENC_ISO))
    If RdTag.Comment(28) = 0 And Not RdTag.Comment(29) = 0 Then
        'ID3v1.1 tag: byte 28 of the comment tag is zeroed, byte 29 is the track number.
        outTag.Comment = Trim$(Data2String(VarPtr(RdTag.Comment(0)), 28, ENC_ISO))
        outTag.TrackNr = RdTag.Comment(29)
    Else
        'ID3v1.0 tag: contains no track number.
        outTag.Comment = Trim$(Data2String(VarPtr(RdTag.Comment(0)), 30, ENC_ISO))
    End If
    outTag.Genre = GetGenreName(RdTag.Genre)
    outTag.SongYear = Trim$(Data2String(VarPtr(RdTag.SongYear(0)), 4, ENC_ISO))

    ReadID3v1 = True
End If
End Function

Public Function WriteID3v1(ByVal FileName As String, ByRef inTag As ID3Tag, Optional ByVal bIgnoreReadOnly As Boolean = True) As Boolean
Dim fh As Long, fsize As Long
Dim rdBuf(2) As Byte, WrTag As v1Tag

If bIgnoreReadOnly Then IgnoreReadOnly FileName
fh = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
If fh = INVALID_HANDLE_VALUE Then Exit Function
fsize = GetFileSize(fh, 0)

'Check if there's already an ID3v1 tag present.
'The ID3v1 tag is located 128 bytes from the end of the file, so look at this position.
If fsize >= 128 Then
    SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
    ReadFile fh, rdBuf(0), 3, 0, ByVal 0
    If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "TAG" Then
        'An ID3v1 tag is present.
        SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
    Else
        'An ID3v1 tag isn't present.
        SetFilePointer fh, fsize, 0, FILE_BEGIN
    End If
Else
    'An ID3v1 tag isn't present.
    SetFilePointer fh, fsize, 0, FILE_BEGIN
End If

'Build up writing tag.
String2Data "TAG", 3, VarPtr(WrTag.Identifier(0)), False
String2Data inTag.Artist, 30, VarPtr(WrTag.Artist(0)), False
String2Data inTag.Album, 30, VarPtr(WrTag.Album(0)), False
String2Data inTag.Title, 30, VarPtr(WrTag.Title(0)), False
String2Data Replace$(inTag.Comment, vbNewLine, " "), 28, VarPtr(WrTag.Comment(0)), False

WrTag.Comment(28) = 0
WrTag.Comment(29) = 0

If IsNumeric(inTag.TrackNr) And Len(Trim$(inTag.TrackNr)) <= 3 Then
    If CLng(inTag.TrackNr) > 0 And CLng(inTag.TrackNr) <= 255 Then
        WrTag.Comment(29) = CByte(inTag.TrackNr)
    End If
End If

WrTag.Genre = GetGenreIndex(inTag.Genre)
String2Data inTag.SongYear, 4, VarPtr(WrTag.SongYear(0)), False

'Write the tag to file.
WriteFile fh, WrTag, Len(WrTag), 0, ByVal 0

CloseHandle fh
WriteID3v1 = True
End Function

Public Function DeleteID3v1(ByVal FileName As String, Optional ByVal bIgnoreReadOnly As Boolean = True) As Boolean
Dim fh As Long, fsize As Long
Dim rdBuf(2) As Byte

If bIgnoreReadOnly Then IgnoreReadOnly FileName
fh = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)

If fh = INVALID_HANDLE_VALUE Then Exit Function
fsize = GetFileSize(fh, 0)

If fsize < 128 Then
    CloseHandle fh
    Exit Function
End If

'Check if there's an ID3v1 tag present.
'The ID3v1 tag is located 128 bytes from the end of the file, so look at this position.
SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
ReadFile fh, rdBuf(0), 3, 0, ByVal 0

If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "TAG" Then
    'An ID3v1 tag is present, remove it.
    SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
    SetEndOfFile fh
End If

CloseHandle fh
DeleteID3v1 = True
End Function
Public Function ReadID3v2(ByVal FileName As String, ByRef outTag As ID3Tag) As Boolean
Dim fh As Long, fp As Long, TagFooter As v2TagHeader
Dim rdBuf(3) As Byte

fh = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
If fh = INVALID_HANDLE_VALUE Then Exit Function

'Search for an ID3v2 tag.
Do
    SetFilePointer fh, fp, 0, FILE_BEGIN
    ReadFile fh, rdBuf(0), 3, 0, ByVal 0
    If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "ID3" Then
        'We've got an ID3 v2 tag.
        fp = ReadV2Tag(fh, fp, outTag, False)
        ReadID3v2 = True
    Else
        fp = -1
    End If
Loop While Not fp = -1

If Not ReadID3v2 Then
    'No tag found at the beginning of the file.
    'Version 2.4 allows the ID3 tag to be appended to the audio data. So, look if we can
    'find such a tag, identified by a footer. According to id3.org, the ID3v2 footer should
    'be added before other tagging systems, therefore prepended to an possibly existing ID3v1 tag.

    fp = GetFileSize(fh, 0)
    SetFilePointer fh, fp - 128, 0, FILE_BEGIN
    ReadFile fh, rdBuf(0), 3, 0, ByVal 0
    
    If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "TAG" Then
        'Look before the ID3 v1 tag.
        fp = fp - 128 - Len(TagFooter)
    Else
        'Look at the end of the file.
        fp = fp - Len(TagFooter)
    End If
    If fp < 0 Then fp = 0

    SetFilePointer fh, fp, 0, FILE_BEGIN
    ReadFile fh, TagFooter, Len(TagFooter), 0, ByVal 0
    
    If Data2String(VarPtr(TagFooter.Identifier(0)), 3, ENC_ISO) = "3DI" Then
        'Got a footer. Calculate the header position from the size.
        fp = fp - Data2Long(VarPtr(TagFooter.Size(0)), True) - Len(TagFooter)
        If Not fp < 0 Then
            ReadV2Tag fh, fp, outTag, False
            ReadID3v2 = True
        End If
    End If
End If
CloseHandle fh
End Function
Private Function ReadV2Tag(ByVal fh As Long, ByVal fp As Long, ByRef RdTag As ID3Tag, ByVal bToMem As Boolean) As Long
Dim TagSizes As v2TagSizes, arrp As Long
Dim TagHeader As v2TagHeader, TagHeaderFlags As v2TagHeaderFlags
Dim ExtHeader23 As v2_3ExtHeader, ExtHeader24 As v2_4ExtHeader
Dim FraHeader22 As v2_2FrameHeader, FraHeader234 As v2_34FrameHeader
Dim FraDataSize As Long, FraUnsynchDecr As Long, fraID As String, FraText As String
Dim FraOffset As Long, bSkipFra As Boolean
Dim TagData() As Byte, SizeBuf(3) As Byte


ReadV2Tag = -1

'------------------------------------------------------
'********************* TAG HEADER *********************
'------------------------------------------------------

SetFilePointer fh, fp, 0, FILE_BEGIN
ReadFile fh, TagHeader, Len(TagHeader), 0, ByVal 0
If TagHeader.Version(0) < 2 Or TagHeader.Version(0) > 4 Then Exit Function
'The size stored in the header excludes itself, and excludes the footer (if present).
TagSizes.ReadSize = Data2Long(VarPtr(TagHeader.Size(0)), True)

'v 2.2 flags: %ab000000 a = unsynch, b = compression (not implemented, ignored)
'v 2.3 flags: %abc00000 a = unsynch, b = extended header, c = experimental (ignored)
'v 2.4 flags: %abcd0000 a = unsynch, b = extended header, c = experimental (ignored), d = footer present
If TagHeader.flags And &H80& Then TagHeaderFlags.bUnsynch = True
If TagHeader.Version(0) >= 3 Then
    If TagHeader.flags And &H40& Then TagHeaderFlags.bExtHeader = True
End If
If TagHeader.Version(0) = 4 Then
    If TagHeader.flags And &H10& Then
        TagHeaderFlags.bFooter = True
        TagSizes.FooterSize = Len(TagHeader)
        'Add the size of the footer (which is the same size than the header) to the read size.
        'Padding must not exist when a footer is present.
        TagSizes.ReadSize = TagSizes.ReadSize + TagSizes.FooterSize
    End If
End If

'Read in the rest of the tag.
ReDim TagData(TagSizes.ReadSize - 1)
ReadFile fh, TagData(0), TagSizes.ReadSize, 0, ByVal 0

'Perform de-unsynchoronisation of the whole read tag, if nescessary.
'In version 2.2 and 2.3, the unsynch is done for all data after the header if the flag is set.
'In version 2.4, the unsynch is applied on a per frame base.
If TagHeader.Version(0) <= 3 And TagHeaderFlags.bUnsynch Then
    TagSizes.ReadSize = TagSizes.ReadSize - DeUnsynch(VarPtr(TagData(0)), TagSizes.ReadSize)
End If


'---------------------------------------------------------------
'********************* EXTENDED TAG HEADER *********************
'---------------------------------------------------------------

If TagHeaderFlags.bExtHeader Then
    If TagHeader.Version(0) = 3 Then
        'Ignore the v 2.3 extended header.
        CopyMemory ExtHeader23, TagData(0), Len(ExtHeader23)
        TagSizes.ExtHeaderSize = Data2Long(VarPtr(ExtHeader23.Size(0)), False) + 4
    ElseIf TagHeader.Version(0) = 4 Then
        'Read the extended header update flag.
        CopyMemory ExtHeader24, TagData(0), Len(ExtHeader24)
        TagSizes.ExtHeaderSize = Data2Long(VarPtr(ExtHeader24.Size(0)), True)
        If ExtHeader24.ExtendedFlags And &H40& Then TagHeaderFlags.bUpdateTag = True
    End If
End If

If Not TagHeaderFlags.bUpdateTag Then
    'Reset output tag.
    ResetTag RdTag
End If


'--------------------------------------------------
'********************* FRAMES *********************
'--------------------------------------------------

arrp = TagSizes.ExtHeaderSize
Do
    If TagHeader.Version(0) = 2 Then

        'Check if there's enough space to read another tag header. We don't have to take
        'care of a footer here because v 2.2 doesn't support footers.
        If arrp < 0 Or arrp > TagSizes.ReadSize - Len(FraHeader22) Then Exit Do

        'Get the frame header.
        CopyMemory FraHeader22, TagData(arrp), Len(FraHeader22)
        arrp = arrp + Len(FraHeader22)

        'Read out the frame ID.
        fraID = Data2String(VarPtr(FraHeader22.Ident(0)), 3, ENC_ISO, False)
        If fraID = String$(3, Chr(0)) Then Exit Do

        'Read out the frame data size.
        CopyMemory SizeBuf(1), FraHeader22.Size(0), 3
        FraDataSize = Data2Long(VarPtr(SizeBuf(0)), False)
        If arrp > TagSizes.ReadSize - FraDataSize Then Exit Do

        'Read out the frame itself.
        If bToMem Then
            'CRM: Encrypted meta frame. Only supported in v 2.2, not converted.
            If Not fraID = "CRM" And FraDataSize > 0 Then
                'Store this frame in the memory buffer, we only need the data to update
                'a currently existing tag.
                ReDim Preserve mMemFrames(mMemFraCnt)
                With mMemFrames(mMemFraCnt)
                    .Id = fraID
                    .Size = FraDataSize
                    ReDim .Data(.Size - 1)
                    CopyMemory .Data(0), TagData(arrp), .Size
                End With
                mMemFraCnt = mMemFraCnt + 1
                ConvertV23 mMemFraCnt - 1
            End If
            FraText = ""
        Else
            FraText = ReadFrame(VarPtr(TagData(arrp)), fraID, FraDataSize)
        End If

    ElseIf TagHeader.Version(0) >= 3 Then

        'Check if there's enough space to read another tag header.
        If arrp < 0 Or arrp > TagSizes.ReadSize - TagSizes.FooterSize - Len(FraHeader234) Then Exit Do

        'Get the frame header.
        CopyMemory FraHeader234, TagData(arrp), Len(FraHeader234)
        arrp = arrp + Len(FraHeader234)

        'Read out the frame ID.
        fraID = Data2String(VarPtr(FraHeader234.Ident(0)), 4, ENC_ISO, False)
        If fraID = String$(4, Chr(0)) Then
            'Got the padding. Frame parsing is finished now.
            Exit Do
        End If

        'Read out the frame data size.
        'v 2.3: No synchsafe integer.
        'v 2.4: Synchsafe integer. However, there are sloppy programs which code version 2.4 with
        'non-synchsafe integers. Do a test therefore, by checking if the next frame is a valid one.
        FraDataSize = Data2Long(VarPtr(FraHeader234.Size(0)), IIf(TagHeader.Version(0) = 3, False, True))
        If arrp > TagSizes.ReadSize - TagSizes.FooterSize - FraDataSize Then Exit Do
        If arrp <= TagSizes.ReadSize - TagSizes.FooterSize - FraDataSize - 4 Then
            'Check if the next frame would be valid.
            If Not IsValidFrameName(VarPtr(TagData(arrp + FraDataSize))) Then
                'Try to decode the number as non-synchsafe.
                FraDataSize = Data2Long(VarPtr(FraHeader234.Size(0)), False)
                If arrp > TagSizes.ReadSize - TagSizes.FooterSize - FraDataSize Then Exit Do
            End If
        End If

        'Read out the frame flags.
        FraOffset = 0
        bSkipFra = False
        FraUnsynchDecr = 0
        If TagHeader.Version(0) = 3 Then
            'v 2.3 frame header flags: %abc00000 %ijk00000
            'abc flags are ignored. i = compression, j = encryption, k = grouping
            If FraHeader234.flags(1) And &HC0& Then
                'Either compressed or encrypted frame. Skip it.
                bSkipFra = True
            End If
            If FraHeader234.flags(1) And &H20& Then
                'This frame belongs to a group. I don't care about that, but we have to
                'consider about the 1 byte added to the header which defines the group.
                FraOffset = 1
            End If
        Else
            'v 2.4 frame header flags: %0abc0000 %0h00kmnp
            'abc flags are ignored. h = grouping, k = compression, m = encryption,
            'n = unsynch, p = data length indicator
            If FraHeader234.flags(1) And &H40& Then
                'See comment above about grouping.
                FraOffset = 1
            End If
            If FraHeader234.flags(1) And &HC& Then
                'Either compressed or encrypted frame. Skip it.
                bSkipFra = True
            End If
            If (FraHeader234.flags(1) And &H2&) Or TagHeaderFlags.bUnsynch Then
                If Not bSkipFra Then
                    'Unsynched frame. DeUnsynch it now.
                    FraUnsynchDecr = DeUnsynch(VarPtr(TagData(arrp)), FraDataSize)
                End If
            End If
            If FraHeader234.flags(1) And &H1& Then
                'Data length indicator. Is ignored, but add 4 bytes offset, because it
                'appends a synchsafe integer to the header.
                FraOffset = FraOffset + 4
            End If
        End If

        'Read out the frame itself.
        If Not bSkipFra Then
            'Skipped frames also are NOT added to the memory, because the new frames
            'are written with a frame header that doesn't include the flags for
            'compression or encryption.
            If bToMem And Not fraID = "SEEK" Then
                'Store this frame in the memory buffer, we only need the data to update
                'a currently existing tag. The v 2.4 seek frame isn't copied to memory,
                'because the new written file is going to have different offsets anyway.
                'However, read out the frame to ensure processing of the appended tag.
                If FraDataSize - FraOffset - FraUnsynchDecr > 0 Then
                    ReDim Preserve mMemFrames(mMemFraCnt)
                    With mMemFrames(mMemFraCnt)
                        .Id = fraID
                        .Size = FraDataSize - FraOffset - FraUnsynchDecr
                        ReDim .Data(.Size - 1)
                        CopyMemory .Data(0), TagData(arrp + FraOffset), .Size
                    End With
                    mMemFraCnt = mMemFraCnt + 1
                    ConvertV23 mMemFraCnt - 1
                End If
                FraText = ""
            Else
                FraText = ReadFrame(VarPtr(TagData(arrp + FraOffset)), fraID, FraDataSize - FraOffset - FraUnsynchDecr)
            End If
        Else
            FraText = ""
        End If
    End If

    'Increase the array pointer.
    arrp = arrp + FraDataSize

    'Write the tags into the structure.
    If fraID = "TP1" Or fraID = "TPE1" Then
        RdTag.Artist = FraText
    ElseIf fraID = "TAL" Or fraID = "TALB" Then
        RdTag.Album = FraText
    ElseIf fraID = "TT2" Or fraID = "TIT2" Then
        RdTag.Title = FraText
    ElseIf fraID = "TRK" Or fraID = "TRCK" Then
        RdTag.TrackNr = FraText
    ElseIf fraID = "TDRC" Then
        'The v 2.4 timestamp frame includes the release year.
        RdTag.SongYear = Left$(FraText, 4)
    ElseIf fraID = "TYE" Or fraID = "TYER" Then
        RdTag.SongYear = FraText
    ElseIf fraID = "TCO" Or fraID = "TCON" Then
        RdTag.Genre = Frame2Genre(FraText)
    ElseIf fraID = "COM" Or fraID = "COM " Or fraID = "COMM" Then
        RdTag.Comment = FraText
    ElseIf fraID = "TEN" Or fraID = "TENC" Then
        RdTag.EncodedBy = FraText
    ElseIf fraID = "WXX" Or fraID = "WXXX" Then
        RdTag.LinkTo = FraText
    ElseIf fraID = "TCR" Or fraID = "TCOP" Then
        RdTag.Copyright = FraText
    ElseIf fraID = "TOA" Or fraID = "TOPE" Then
        RdTag.OrigArtist = FraText
    ElseIf fraID = "TCM" Or fraID = "TCOM" Then
        RdTag.Composer = FraText
    ElseIf fraID = "SEEK" Then
        If IsNumeric(FraText) Then
            'FraText actually has to be numeric, because its value is generated from Data2Long.
            ReadV2Tag = fp + Len(TagHeader) + TagSizes.ReadSize + CLng(FraText)
        End If
    End If
Loop
End Function

Private Function ReadFrame(ByVal pData As Long, ByVal FrameID As String, ByVal FrameSize As Long) As String
Dim EncFormat As Byte, SkipOffset As Long

If FrameSize = 0 Then Exit Function

CopyMemory EncFormat, ByVal pData, 1
If Left$(FrameID, 1) = "T" Then
    'Text frame.
    ReadFrame = Data2String(pData + 1, FrameSize - 1, EncFormat)
ElseIf FrameID = "COMM" Or FrameID = "COM " Or FrameID = "COM" Then
    'Comment frame.
    'Skip the 3 bytes language identifier.
    SkipOffset = 4
    'Skip the content description.
    ReadFrame = Data2String(pData + SkipOffset, FrameSize - SkipOffset, EncFormat)
    If EncFormat = ENC_ISO Or EncFormat = ENC_UNICODE_UTF8 Then
        SkipOffset = SkipOffset + Len(ReadFrame) + 1
    Else
        SkipOffset = SkipOffset + Len(ReadFrame) * 2 + 2
    End If
    'Read the actual text.
    ReadFrame = Data2String(pData + SkipOffset, FrameSize - SkipOffset, EncFormat)
ElseIf FrameID = "WXXX" Or FrameID = "WXX" Then
    'Link frame.
    SkipOffset = 1
    ReadFrame = Data2String(pData + SkipOffset, FrameSize - SkipOffset, EncFormat)
    If EncFormat = ENC_ISO Or EncFormat = ENC_UNICODE_UTF8 Then
        SkipOffset = SkipOffset + Len(ReadFrame) + 1
    Else
        SkipOffset = SkipOffset + Len(ReadFrame) * 2 + 2
    End If
    ReadFrame = Data2String(pData + SkipOffset, FrameSize - SkipOffset, EncFormat)
ElseIf FrameID = "SEEK" Then
    If FrameSize < 4 Then Exit Function
    'Got a seek frame (v 2.4 only). It points to the next tag in the file.
    'According to id3.org, it seems that this value isn't synchsafe... ??
    ReadFrame = Data2Long(pData, False)
End If
End Function

Private Function IsValidFrameName(ByVal pData As Long) As Boolean
Dim i As Long, curData As Byte

For i = 0 To 3
    CopyMemory curData, ByVal pData + i, 1
    'Accept $00 $00 $00 $00 as a valid frame name, because this would be the start of the pattern.
    If Not curData = 0 And Not (curData >= Asc("A") And curData <= Asc("Z")) And Not (curData >= Asc("0") And curData <= Asc("9")) Then
        Exit Function
    End If
Next i

IsValidFrameName = True
End Function
Private Function DeUnsynch(ByVal pData As Long, ByVal length As Long) As Long
Dim i As Long, subIdx As Long
Dim curData(1) As Byte

Do
    'The following line performs pData(subIdx) = pData(i)
    CopyMemory ByVal pData + subIdx, ByVal pData + i, 1
    subIdx = subIdx + 1
    If i <= length - 2 Then
        CopyMemory curData(0), ByVal pData + i, 2
        If curData(0) = &HFF& And curData(1) = &H0& Then
            'Jump over the &H0& byte.
            i = i + 2
            DeUnsynch = DeUnsynch + 1
            If i = length Then Exit Do
        Else
            i = i + 1
        End If
    Else
        Exit Do
    End If
Loop
End Function
Private Sub ConvertV23(ByVal memIdx As Long)
Dim i As Long
Dim sPicType As String, expSize As Long

If memIdx < 0 Or memIdx > mMemFraCnt - 1 Then Exit Sub
With mMemFrames(memIdx)
    If .Id = "PIC" Then
        'Convert the picture type desc from v 2.2 to 2.3 (and 2.4)
        sPicType = UCase$(Data2String(VarPtr(.Data(1)), 3, ENC_ISO))
        If sPicType = "PNG" Then
            sPicType = "image/png"
        ElseIf sPicType = "JPG" Then
            sPicType = "image/jpeg"
        ElseIf Not sPicType = "-->" Then
            sPicType = "image/" & LCase$(sPicType)
        End If
        'Expand the data.
        expSize = Len(sPicType) - 2
        If expSize <= 0 Then Exit Sub
        .Size = .Size + expSize
        ReDim Preserve .Data(.Size - 1)
        'Copy the data back, because we need more space at the beginning for the string.
        For i = .Size - 1 To 4 + expSize Step -1
            .Data(i) = .Data(i - expSize)
        Next i
        'Write the string into the data field.
        String2Data sPicType, 0, VarPtr(.Data(1)), True
    ElseIf .Id = "LNK" Then
        'Convert the link frame from v 2.2 to 2.3 (and 2.4)
        sPicType = Data2String(VarPtr(.Data(0)), 3, ENC_ISO)
        .Size = .Size + 1
        ReDim Preserve .Data(.Size - 1)
        For i = .Size - 1 To 4 Step -1
            .Data(i) = .Data(i - 1)
        Next i
        String2Data GetNewFrameID(sPicType), 4, VarPtr(.Data(0)), False
    ElseIf .Id = "TDOR" Then
        'Convert the TDOR timestamp from 2.4 to 2.3, simply by using only the first
        '4 bytes, identifying the year.
        .Size = 4
    End If
    .Id = GetNewFrameID(.Id)
End With
End Sub

Private Function GetNewFrameID(ByVal sID As String) As String
If sID = "UFI" Then
    sID = "UFID"
ElseIf sID = "TT1" Then
    sID = "TIT1"
ElseIf sID = "TT2" Then
    sID = "TIT2"
ElseIf sID = "TT3" Then
    sID = "TIT3"
ElseIf sID = "TP1" Then
    sID = "TPE1"
ElseIf sID = "TP2" Then
    sID = "TPE2"
ElseIf sID = "TP3" Then
    sID = "TPE3"
ElseIf sID = "TP4" Then
    sID = "TPE4"
ElseIf sID = "TCM" Then
    sID = "TCOM"
ElseIf sID = "TXT" Then
    sID = "TEXT"
ElseIf sID = "TLA" Then
    sID = "TLAN"
ElseIf sID = "TCO" Then
    sID = "TCON"
ElseIf sID = "TAL" Then
    sID = "TALB"
ElseIf sID = "TPA" Then
    sID = "TPOS"
ElseIf sID = "TRK" Then
    sID = "TRCK"
ElseIf sID = "TRC" Then
    sID = "TSRC"
ElseIf sID = "TYE" Then
    sID = "TYER"
ElseIf sID = "TDA" Then
    sID = "TDAT"
ElseIf sID = "TIM" Then
    sID = "TIME"
ElseIf sID = "TRD" Then
    sID = "TRDA"
ElseIf sID = "TMT" Then
    sID = "TMED"
ElseIf sID = "TFT" Then
    sID = "TFLT"
ElseIf sID = "TBP" Then
    sID = "TBPM"
ElseIf sID = "TCR" Then
    sID = "TCOP"
ElseIf sID = "TPB" Then
    sID = "TPUB"
ElseIf sID = "TEN" Then
    sID = "TENC"
ElseIf sID = "TSS" Then
    sID = "TSSE"
ElseIf sID = "TOF" Then
    sID = "TOFN"
ElseIf sID = "TLE" Then
    sID = "TLEN"
ElseIf sID = "TSI" Then
    sID = "TSIZ"
ElseIf sID = "TDY" Then
    sID = "TDLY"
ElseIf sID = "TKE" Then
    sID = "TKEY"
ElseIf sID = "TOT" Then
    sID = "TOAL"
ElseIf sID = "TOL" Then
    sID = "TOLY"
ElseIf sID = "TOR" Or sID = "TDOR" Then
    sID = "TORY"
ElseIf sID = "TXX" Then
    sID = "TXXX"
ElseIf sID = "WAF" Then
    sID = "WOAF"
ElseIf sID = "WAR" Then
    sID = "WOAR"
ElseIf sID = "WAS" Then
    sID = "WOAS"
ElseIf sID = "WCM" Then
    sID = "WCOM"
ElseIf sID = "WCP" Then
    sID = "WCOP"
ElseIf sID = "WPB" Then
    sID = "WPUB"
ElseIf sID = "WXX" Then
    sID = "WXXX"
ElseIf sID = "IPL" Or sID = "TIPL" Then
    sID = "IPLS"
ElseIf sID = "MCI" Then
    sID = "MCDI"
ElseIf sID = "ETC" Then
    sID = "ETCO"
ElseIf sID = "MLL" Then
    sID = "MLLT"
ElseIf sID = "STC" Then
    sID = "SYTC"
ElseIf sID = "ULT" Then
    sID = "USLT"
ElseIf sID = "SLT" Then
    sID = "SYLT"
ElseIf sID = "COM" Then
    sID = "COMM"
ElseIf sID = "RVA" Then
    sID = "RVAD"
ElseIf sID = "EQU" Then
    sID = "EQUA"
ElseIf sID = "REV" Then
    sID = "RVRB"
ElseIf sID = "PIC" Then
    sID = "APIC"
ElseIf sID = "GEO" Then
    sID = "GEOB"
ElseIf sID = "CNT" Then
    sID = "PCNT"
ElseIf sID = "POP" Then
    sID = "POPM"
ElseIf sID = "BUF" Then
    sID = "RBUF"
ElseIf sID = "CRA" Then
    sID = "AENC"
ElseIf sID = "LNK" Then
    sID = "LINK"
End If

GetNewFrameID = sID
End Function
Public Function WriteID3v2(ByVal FileName As String, ByRef inTag As ID3Tag, ByVal Version As ID3v2_Version, ByVal bClearExistingTag As Boolean, Optional ByVal bIgnoreReadOnly As Boolean = True, Optional ByVal pBuffer As Long = 0, Optional ByVal BufferSize As Long = 0) As Boolean
Dim i As Long, fh As Long, fp As Long, fsize As Long
Dim TagHeader As v2TagHeader, ExistSize As Long, arrp As Long
Dim WrData() As Byte, WrSize As Long, WrBuf() As Byte, BuildTag As ID3Tag
Dim tmpHeader As v2_34FrameHeader, rdBuf(3) As Byte

If bIgnoreReadOnly Then IgnoreReadOnly FileName
fh = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
If fh = INVALID_HANDLE_VALUE Then Exit Function

'Never mind on possibly existing appended tags or footer tags. Just leave them in the file.
ExistSize = GetTagSize(fh, 0)

'Read the existing tag into memory if it shouldn't be cleared.
mMemFraCnt = 0
If Not ExistSize = 0 And Not bClearExistingTag Then
    Do
        SetFilePointer fh, fp, 0, FILE_BEGIN
        ReadFile fh, rdBuf(0), 3, 0, ByVal 0
        If Data2String(VarPtr(rdBuf(0)), 3, ENC_ISO) = "ID3" Then
            'We've got an ID3 v2 tag. Write it temporary in the BuildTag variable.
            fp = ReadV2Tag(fh, fp, BuildTag, True)
        Else
            fp = -1
        End If
    Loop While Not fp = -1
End If

BuildTag = inTag
BuildTag.Genre = Genre2Frame(BuildTag.Genre)

'Determine required size for the new tag.
WrSize = Len(TagHeader)
With BuildTag
    WrSize = WrSize + GetFrameSize("TPE1", .Artist, True)
    WrSize = WrSize + GetFrameSize("TALB", .Album, True)
    WrSize = WrSize + GetFrameSize("TIT2", .Title, True)
    WrSize = WrSize + GetFrameSize(IIf(Version = VERSION_2_3, "TYER", "TDRC"), .SongYear, True)
    WrSize = WrSize + GetFrameSize("TCON", .Genre, True)
    WrSize = WrSize + GetFrameSize("COMM", .Comment, True)
    WrSize = WrSize + GetFrameSize("TRCK", .TrackNr, True)
    WrSize = WrSize + GetFrameSize("TENC", .EncodedBy, True)
    WrSize = WrSize + GetFrameSize("WXXX", .LinkTo, True)
    WrSize = WrSize + GetFrameSize("TCOP", .Copyright, True)
    WrSize = WrSize + GetFrameSize("TCOM", .Composer, True)
    WrSize = WrSize + GetFrameSize("TOPE", .OrigArtist, True)
End With
'Add the size of the existing frames which should be preserved.
For i = 0 To mMemFraCnt - 1
    If IsToStore(mMemFrames(i).Id) Then
        WrSize = WrSize + mMemFrames(i).Size + Len(tmpHeader)
    End If
Next i

If WrSize <= ExistSize Then
    'Overwrite the existing tag completely.
    WrSize = ExistSize
Else
    'We have to rewrite the file anyway.
    'Add padding, and not too less. A few KB don't matter in a 3+ MB file, and the file won't
    'have to be rewritten soon.
    If WrSize Mod 2048 < 1024 Then
        WrSize = (WrSize \ 2048) * 2048 + 2048
    Else
        WrSize = (WrSize \ 2048) * 2048 + 4096
    End If
    WrSize = WrSize + Len(TagHeader)
End If
ReDim WrData(WrSize - 1)


'***** Build up tag header *****
String2Data "ID3", 3, VarPtr(TagHeader.Identifier(0)), False
TagHeader.Version(0) = IIf(Version = VERSION_2_3, 3, 4)
Long2Data WrSize - Len(TagHeader), True, VarPtr(TagHeader.Size(0))
CopyMemory WrData(0), TagHeader, Len(TagHeader)
arrp = Len(TagHeader)

'***** No extended header is used, since it even seems that many programs don't support them *****

'***** Start writing the frame data into the array *****
With BuildTag
    'Note: we don't use VarPtr(WrData(arrp)) here because if arrp == <array size>, the function
    'call would fail, although there is no more data written to the array.
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TPE1", .Artist, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TALB", .Album, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TIT2", .Title, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, IIf(Version = VERSION_2_3, "TYER", "TDRC"), .SongYear, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TCON", .Genre, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "COMM", .Comment, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TRCK", .TrackNr, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TENC", .EncodedBy, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "WXXX", .LinkTo, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TCOP", .Copyright, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TCOM", .Composer, Version)
    arrp = arrp + WriteFrame(VarPtr(WrData(0)) + arrp, "TOPE", .OrigArtist, Version)
End With
For i = 0 To mMemFraCnt - 1
    If IsToStore(mMemFrames(i).Id) Then
        arrp = arrp + WriteMemFrame(VarPtr(WrData(0)) + arrp, i, Version)
    End If
Next i

'***** Padding is added "automatically" because the WrData() array is zeroed from VB *****


'Write the data out to the file.
If WrSize <= ExistSize Then
    'Here we go, just overwrite the existing tag.
    SetFilePointer fh, 0, 0, FILE_BEGIN
    WriteFile fh, WrData(0), WrSize, 0, ByVal 0
Else
    'We need to rewrite the entire file...
    fsize = GetFileSize(fh, 0)
    If pBuffer = 0 Or BufferSize < fsize - ExistSize Then
        'Allocate a dynamic buffer.
        ReDim WrBuf(fsize - ExistSize - 1)
        pBuffer = VarPtr(WrBuf(0))
    End If

    'Use the buffer to store the audio data of the file.
    SetFilePointer fh, ExistSize, 0, FILE_BEGIN
    ReadFile fh, ByVal pBuffer, fsize - ExistSize, 0, ByVal 0
    'Compute the new file size and expand the existing file.
    fsize = WrSize + fsize - ExistSize
    SetFilePointer fh, fsize, 0, FILE_BEGIN
    SetEndOfFile fh
    'Write the new tag to the file.
    SetFilePointer fh, 0, 0, FILE_BEGIN
    WriteFile fh, WrData(0), WrSize, 0, ByVal 0
    'Append the audio data again.
    WriteFile fh, ByVal pBuffer, fsize - WrSize, 0, ByVal 0
End If

CloseHandle fh
WriteID3v2 = True
End Function
Private Function IsToStore(ByVal fraID As String) As Boolean
IsToStore = True

If fraID = "TPE1" Or fraID = "TALB" Or fraID = "TIT2" Then
    IsToStore = False
ElseIf fraID = "TYER" Or fraID = "TDRC" Or fraID = "TCON" Or fraID = "COMM" Then
    IsToStore = False
ElseIf fraID = "TRCK" Or fraID = "TENC" Or fraID = "WXXX" Or fraID = "TCOP" Then
    IsToStore = False
ElseIf fraID = "TCOM" Or fraID = "TOPE" Then
    IsToStore = False
End If
End Function

Private Function WriteFrame(ByVal pData As Long, ByVal FrameID As String, ByVal strData As String, ByVal Version As ID3v2_Version) As Long
Dim i As Long, fraHeader As v2_34FrameHeader
Dim EncFormat As Byte, WrOffset As Long

If Trim$(strData) = "" Then Exit Function

'Create the frame header.
String2Data FrameID, 4, VarPtr(fraHeader.Ident(0)), False
WriteFrame = GetFrameSize(FrameID, strData, False)
Long2Data WriteFrame, IIf(Version = VERSION_2_3, False, True), VarPtr(fraHeader.Size(0))
CopyMemory ByVal pData, fraHeader, Len(fraHeader)
WrOffset = Len(fraHeader)

'Add the size of the header to the return value.
WriteFrame = WriteFrame + Len(fraHeader)

'Create the frame data.
If Left$(FrameID, 1) = "T" Then
    'Text frame.
    'EncFormat = 0 = ISO encoding
    CopyMemory ByVal pData + WrOffset, EncFormat, 1
    WrOffset = WrOffset + 1
    String2Data strData, 0, pData + WrOffset, True
ElseIf FrameID = "COMM" Then
    'Comment frame.
    CopyMemory ByVal pData + WrOffset, EncFormat, 1
    WrOffset = WrOffset + 1
    'English language.
    String2Data "eng", 0, pData + WrOffset, False
    WrOffset = WrOffset + 3
    'Create an empty content description.
    CopyMemory ByVal pData + WrOffset, EncFormat, 1
    WrOffset = WrOffset + 1
    String2Data strData, 0, pData + WrOffset, True
ElseIf FrameID = "WXXX" Then
    'Link frame.
    For i = 0 To 1
        'i=0: encoding format, i=1: empty description string
        CopyMemory ByVal pData + WrOffset, EncFormat, 1
        WrOffset = WrOffset + 1
    Next i
    String2Data strData, 0, pData + WrOffset, True
Else
    Exit Function
End If
End Function

Private Function WriteMemFrame(ByVal pData As Long, ByVal fraIdx As Long, ByVal Version As ID3v2_Version) As Long
Dim i As Long, fraHeader As v2_34FrameHeader
Dim EncFormat As Byte, WrOffset As Long

If fraIdx < 0 Or fraIdx > mMemFraCnt - 1 Then Exit Function
With mMemFrames(fraIdx)
    If Version = VERSION_2_4 Then
        If .Id = "TORY" Then
            .Id = "TDOR"
        ElseIf .Id = "IPLS" Then
            .Id = "TIPL"
        End If
    End If
    'Create the frame header.
    String2Data .Id, 4, VarPtr(fraHeader.Ident(0)), False
    WriteMemFrame = .Size
    Long2Data WriteMemFrame, IIf(Version = VERSION_2_3, False, True), VarPtr(fraHeader.Size(0))
    CopyMemory ByVal pData, fraHeader, Len(fraHeader)
    WrOffset = Len(fraHeader)

    'Add the size of the header to the return value.
    WriteMemFrame = WriteMemFrame + Len(fraHeader)

    'Create the frame data.
    CopyMemory ByVal pData + WrOffset, .Data(0), .Size
End With
End Function

Private Function GetFrameSize(ByVal FrameID As String, ByVal strData As String, ByVal bIncludeHeader As Boolean) As Long
Dim fraHeader As v2_34FrameHeader

If Trim$(strData) = "" Then Exit Function

If bIncludeHeader Then GetFrameSize = Len(fraHeader)
If Left$(FrameID, 1) = "T" Then
    'Text frame: 1 b encoding, x b string, 1 b null
    GetFrameSize = GetFrameSize + Len(strData) + 2
ElseIf FrameID = "COMM" Then
    'Comment frame: 1 b encoding, 3 b language, 1 b empty description, x b string, 1 b null
    GetFrameSize = GetFrameSize + Len(strData) + 6
ElseIf FrameID = "WXXX" Then
    'Link frame: 1 b encoding, 1 b empty description, x b string, 1 b null
    GetFrameSize = GetFrameSize + Len(strData) + 3
Else
    GetFrameSize = 0
End If
End Function


'See the WriteID3v2 function for a description on the Buffer parameters.
Public Function DeleteID3v2(ByVal FileName As String, Optional ByVal bIgnoreReadOnly As Boolean = True, Optional ByVal pBuffer As Long = 0, Optional ByVal BufferSize As Long = 0) As Boolean
Dim fh As Long, fsize As Long
Dim TagHeader As v2TagHeader, ExistSize As Long, WrBuf() As Byte

If bIgnoreReadOnly Then IgnoreReadOnly FileName
fh = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
If fh = INVALID_HANDLE_VALUE Then Exit Function

'Never mind on possibly existing appended tags or footer tags. Just leave them in the file.
ExistSize = GetTagSize(fh, 0)
If Not ExistSize = 0 Then
    fsize = GetFileSize(fh, 0)
    
    If pBuffer = 0 Or BufferSize < fsize - ExistSize Then
        'Allocate a dynamic buffer.
        ReDim WrBuf(fsize - ExistSize - 1)
        pBuffer = VarPtr(WrBuf(0))
    End If

    'Use the buffer to store the audio data of the file.
    SetFilePointer fh, ExistSize, 0, FILE_BEGIN
    ReadFile fh, ByVal pBuffer, fsize - ExistSize, 0, ByVal 0
    
    'Compute the new file size and crop the existing file.
    fsize = fsize - ExistSize
    SetFilePointer fh, fsize, 0, FILE_BEGIN
    SetEndOfFile fh
    
    'Write the audio data to the file again.
    SetFilePointer fh, 0, 0, FILE_BEGIN
    WriteFile fh, ByVal pBuffer, fsize, 0, ByVal 0
End If

CloseHandle fh
DeleteID3v2 = True
End Function
Private Function Data2Long(ByVal pData As Long, ByVal bSynchSafe As Boolean) As Long
Dim i As Integer, Data(3) As Byte

CopyMemory Data(0), ByVal pData, 4

For i = 0 To 3
    If Data(i) And &H80& Then bSynchSafe = False
Next i

'Perform left-shifts, done by multiplication with the hex values of 2^n. Finally, bit-or the values.
If bSynchSafe Then
    Data2Long = (Data(0) * &H200000) Or (Data(1) * &H4000&) Or (Data(2) * &H80&) Or Data(3)
Else
    Data2Long = (Data(0) * &H1000000) Or (Data(1) * &H10000) Or (Data(2) * &H100&) Or Data(3)
End If
End Function

Private Function Long2Data(ByVal SrcValue As Long, ByVal bSynchSafe As Boolean, ByVal pData As Long)
Dim Data(3) As Byte

'Mask out the corresponding bits and right-shift them afterwards (done by division).
If bSynchSafe Then
    Data(0) = (SrcValue And &HFE00000) \ &H200000
    Data(1) = (SrcValue And &H1FC000) \ &H4000&
    Data(2) = (SrcValue And &H3F80&) \ &H80&
    Data(3) = SrcValue And &H7F&
Else
    Data(0) = SrcValue \ &H1000000
    Data(1) = (SrcValue And &HFF0000) \ &H10000
    Data(2) = (SrcValue And &HFF00&) \ &H100&
    Data(3) = SrcValue And &HFF&
End If

'Copy back.
CopyMemory ByVal pData, Data(0), 4
End Function

Private Function Data2String(ByVal pData As Long, ByVal length As Long, ByVal EncFormat As v2_StrEncoding, Optional ByVal BreakOnNull As Boolean = True) As String
Dim i As Long, curData As Byte, curSign As String

For i = 0 To length - 1
    CopyMemory curData, ByVal pData + i, 1
    'New lines are represented by &0A& (which is chr(10)) only in ID3 v2 tags.
    'However, many programs seem to code the newline with chr(13) & chr(10),
    'which is the windows default. Therefore, skip chr(13) and change chr(10) to vbNewLines.
    If curData = 13 Then
        curSign = ""
    ElseIf curData = 10 Then
        curSign = vbNewLine
    Else
        curSign = Chr(curData)
    End If
    If EncFormat = ENC_ISO Or EncFormat = ENC_UNICODE_UTF8 Then
        'Clear text, null terminated.
        If curData = 0 And BreakOnNull Then
            Exit Function
        Else
            Data2String = Data2String & curSign
        End If
    ElseIf EncFormat = ENC_UNICODE_UTF16_BOM Then
        'UNICODE text with BOM, double-null terminated.
        If i >= 2 And i Mod 2 = 0 Then
            If curData = 0 And BreakOnNull Then
                Exit Function
            Else
                Data2String = Data2String & curSign
            End If
        End If
    ElseIf EncFormat = ENC_UNICODE_UTF16 Then
        'UNICODE text without BOM, double-null terminated.
        If i Mod 2 = 0 Then
            If curData = 0 And BreakOnNull Then
                Exit Function
            Else
                Data2String = Data2String & curSign
            End If
        End If
    End If
Next i
End Function

Private Function String2Data(ByVal SrcStr As String, ByVal MaxLength As Long, ByVal pData As Long, ByVal bTerminate As Boolean) As Long
Dim i As Long, SrcLen As Long
Dim curAsc As Byte

SrcLen = Len(SrcStr)

If MaxLength <= 0 Then
    MaxLength = SrcLen
ElseIf MaxLength > SrcLen Then
    MaxLength = SrcLen
End If

For i = 0 To MaxLength - 1
    curAsc = Asc(Mid$(SrcStr, i + 1, 1))
    CopyMemory ByVal pData + i, curAsc, 1
Next i

If bTerminate Then
    'Add the null terminator.
    curAsc = 0
    CopyMemory ByVal pData + MaxLength, curAsc, 1
    String2Data = MaxLength + 1
Else
    String2Data = MaxLength
End If
End Function

Private Function Frame2Genre(ByVal strData As String) As String
Dim idxFrom As Long, idxTo As Long
Dim strRef As String

Frame2Genre = strData
idxFrom = InStr(1, strData, "(")
idxTo = InStr(1, strData, ")")
If idxFrom = 0 Or idxTo = 0 Then Exit Function

strRef = Mid$(Frame2Genre, idxFrom + 1, idxTo - idxFrom - 1)
If IsNumeric(strRef) Then
    Frame2Genre = GetGenreName(CLng(strRef))
ElseIf strRef = "RX" Then
    Frame2Genre = "Remix"
ElseIf strRef = "CR" Then
    Frame2Genre = "Cover"
End If
End Function

Private Function Genre2Frame(ByVal strGenre As String) As String
Dim GenreIdx As Long

GenreIdx = GetGenreIndex(strGenre)

If GenreIdx = 255 Then
    Genre2Frame = strGenre
Else
    Genre2Frame = "(" & GenreIdx & ")" & GetGenreName(GenreIdx)
End If
End Function

Public Function GetGenreName(ByVal GenreIndex As Long) As String
If GenreIndex < 0 Or GenreIndex > 147 Then Exit Function

If Not bBuiltGenre Then
    bBuiltGenre = True
    GenreField = Split(cGenreString, "#")
End If

GetGenreName = GenreField(GenreIndex)
End Function

Public Function GetGenreIndex(ByVal GenreName As String) As Long
Dim i As Long

If Not bBuiltGenre Then
    bBuiltGenre = True
    GenreField = Split(cGenreString, "#")
End If

GenreName = LCase$(Trim$(GenreName))

For i = 0 To 147
    If LCase$(GenreField(i)) = GenreName Then
        GetGenreIndex = i
        Exit Function
    End If
Next i

GetGenreIndex = 255
End Function
Public Function ReadMPEGInfo(ByVal FileName As String, ByRef outInfo As MPEGInfo) As Boolean
Dim fh As Long, fp As Long, fsize As Long
Dim fraHeader(3) As Byte, mVers As Long, lVers As Long
Dim opByte As Byte, opLong As Long, VBRHeaderPos(1) As Long
Dim VBRHeader(17) As Byte, VBRIdent As String
Dim SamplesPerFrame As Long, VBRNrFrames As Long, VBRBytes As Long
Dim RdTag As v1Tag, TagHeader As v2TagHeader

fh = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
If fh = INVALID_HANDLE_VALUE Then Exit Function

fp = GetTagSize(fh, 0)
fsize = GetFileSize(fh, 0)
If fsize < 4 Then
    CloseHandle fh
    Exit Function
End If

'Search for the MP3 frame header.
Do
    SetFilePointer fh, fp, 0, FILE_BEGIN
    ReadFile fh, fraHeader(0), 4, 0, ByVal 0
    If fraHeader(0) = &HFF& And (fraHeader(1) And &HE0&) = &HE0& Then
        'Got the frame synchronisation (11 bits set).
        'Read the information from the set bits.
        
        'MPEG version
        opByte = (fraHeader(1) And &H18&) \ &H8&
        If opByte = 0 Then
            outInfo.MPEGVersion = "version 2.5"
            mVers = 3
        ElseIf opByte = 2 Then
            outInfo.MPEGVersion = "version 2"
            mVers = 2
        ElseIf opByte = 3 Then
            outInfo.MPEGVersion = "version 1"
            mVers = 1
        Else
            Exit Do
        End If
        
        'Layer version
        opByte = (fraHeader(1) And &H6&) \ &H2&
        If opByte = 1 Then
            lVers = 3
            SamplesPerFrame = IIf(mVers = 1, 1152, 576)
        ElseIf opByte = 2 Then
            lVers = 2
            SamplesPerFrame = 1152
        ElseIf opByte = 3 Then
            lVers = 1
            SamplesPerFrame = 384
        Else
            Exit Do
        End If
        outInfo.MPEGVersion = outInfo.MPEGVersion & " layer " & lVers
        
        'CRC
        outInfo.HasCRC = IIf(fraHeader(1) And &H1&, False, True)
        
        'Channel mode.
        opByte = (fraHeader(3) And &HC0&) \ &H40&
        
        'The position of an possibly existing VBR header depends on the channel mode.
        If Not opByte = 3 Then
            VBRHeaderPos(0) = fp + 4 + IIf(mVers = 1, 32, 17)
        Else
            VBRHeaderPos(0) = fp + 4 + IIf(mVers = 1, 17, 9)
        End If
        VBRHeaderPos(1) = fp + 4 + 32
        
        'Channel mode.
        If opByte = 0 Then
            outInfo.ChannelMode = "Stereo"
        ElseIf opByte = 1 Then
            outInfo.ChannelMode = "Joint stereo"
        ElseIf opByte = 2 Then
            outInfo.ChannelMode = "Dual channel"
        ElseIf opByte = 3 Then
            outInfo.ChannelMode = "Mono"
        End If
        
        'Frequency
        opByte = (fraHeader(2) And &HC&) \ &H4&
        If opByte = 0 Then
            opLong = 44100
        ElseIf opByte = 1 Then
            opLong = 48000
        ElseIf opByte = 2 Then
            opLong = 32000
        End If
        outInfo.Frequency = opLong / IIf(mVers = 3, 4, mVers)
        
        'Check if there is a VBR header. If present, use it to read out the number of frames.
        SetFilePointer fh, VBRHeaderPos(0), 0, FILE_BEGIN
        ReadFile fh, VBRHeader(0), 16, 0, ByVal 0
        VBRIdent = Data2String(VarPtr(VBRHeader(0)), 4, ENC_ISO)
        If VBRIdent = "Xing" Or VBRIdent = "Info" Then
            'A VBR header is present.
            If VBRHeader(7) And &H1& Then
                'The number of frames is stored.
                VBRNrFrames = Data2Long(VarPtr(VBRHeader(8)), False)
                If VBRHeader(7) And &H2& Then
                    VBRBytes = Data2Long(VarPtr(VBRHeader(12)), False)
                End If
            End If
        End If
        'Check if there is a VBRI header.
        SetFilePointer fh, VBRHeaderPos(1), 0, FILE_BEGIN
        ReadFile fh, VBRHeader(0), 18, 0, ByVal 0
        If Data2String(VarPtr(VBRHeader(0)), 4, ENC_ISO) = "VBRI" Then
            VBRBytes = Data2Long(VarPtr(VBRHeader(10)), False)
            VBRNrFrames = Data2Long(VarPtr(VBRHeader(14)), False)
        End If
        
        If Not VBRBytes = 0 And Not VBRNrFrames = 0 Then
            'VBR bitrate.
            outInfo.HasVBR = True
            outInfo.Bitrate = VBRBytes / VBRNrFrames / SamplesPerFrame / 125 * outInfo.Frequency
            outInfo.length = VBRNrFrames / outInfo.Frequency * SamplesPerFrame
        Else
            'CBR bitrate
            outInfo.HasVBR = False
            opByte = (fraHeader(2) And &HF0&) \ &H10&
            If Not opByte = 0 And Not opByte = 15 Then
                If mVers = 1 And lVers = 1 Then
                    outInfo.Bitrate = opByte * 32
                ElseIf mVers = 1 And lVers = 2 Then
                    If opByte = 1 Then
                        outInfo.Bitrate = 32
                    ElseIf opByte = 2 Then
                        outInfo.Bitrate = 48
                    ElseIf opByte <= 4 Then
                        outInfo.Bitrate = 48 + (opByte - 2) * 8
                    ElseIf opByte <= 8 Then
                        outInfo.Bitrate = 64 + (opByte - 4) * 16
                    ElseIf opByte <= 12 Then
                        outInfo.Bitrate = 128 + (opByte - 8) * 32
                    Else
                        outInfo.Bitrate = 256 + (opByte - 12) * 64
                    End If
                ElseIf mVers = 1 And lVers = 3 Then
                    If opByte = 1 Then
                        outInfo.Bitrate = 32
                    ElseIf opByte <= 5 Then
                        outInfo.Bitrate = 32 + (opByte - 1) * 8
                    ElseIf opByte <= 9 Then
                        outInfo.Bitrate = 64 + (opByte - 5) * 16
                    ElseIf opByte <= 13 Then
                        outInfo.Bitrate = 128 + (opByte - 9) * 32
                    Else
                        outInfo.Bitrate = 320
                    End If
                ElseIf mVers >= 2 And lVers = 1 Then
                    If opByte = 1 Then
                        outInfo.Bitrate = 32
                    ElseIf opByte = 2 Then
                        outInfo.Bitrate = 48
                    ElseIf opByte <= 4 Then
                        outInfo.Bitrate = 48 + (opByte - 2) * 8
                    ElseIf opByte <= 12 Then
                        outInfo.Bitrate = 64 + (opByte - 4) * 16
                    Else
                        outInfo.Bitrate = 192 + (opByte - 12) * 32
                    End If
                Else
                    'mVers >= 2, lVers >= 2
                    If opByte <= 8 Then
                        outInfo.Bitrate = opByte * 8
                    Else
                        outInfo.Bitrate = 64 + (opByte - 8) * 16
                    End If
                End If
            End If
            outInfo.length = (fsize - fp) / (outInfo.Bitrate * 125)
        End If
                    
        'Copyright
        outInfo.IsCopyrighted = IIf(fraHeader(3) And &H8&, True, False)
        
        'Original
        outInfo.IsOriginal = IIf(fraHeader(3) And &H4&, True, False)
        
        'Emphasis
        outInfo.HasEmphasis = IIf(fraHeader(3) And &H3&, True, False)
        
        ReadMPEGInfo = True
        Exit Do
    Else
        fp = fp + 1
        If fp > fsize - 4 Then Exit Do
    End If
Loop

'Read out the tag versions.
'ID3 v1 tag.
outInfo.ID3v1Version = -1
If fsize >= 128 Then
    'The ID3v1 tag is located 128 bytes from the end of the file, so look at this position.
    SetFilePointer fh, fsize - 128, 0, FILE_BEGIN
    ReadFile fh, RdTag, Len(RdTag), 0, ByVal 0
    If Data2String(VarPtr(RdTag.Identifier(0)), 3, ENC_ISO) = "TAG" Then
        If RdTag.Comment(28) = 0 And Not RdTag.Comment(29) = 0 Then
            'ID3v1.1 tag: byte 28 of the comment tag is zeroed, byte 29 is the track number.
            outInfo.ID3v1Version = 1
        Else
            'ID3v1.0 tag: contains no track number.
            outInfo.ID3v1Version = 0
        End If
    End If
End If

'ID3 v2 tag.
outInfo.ID3v2Version = -1
SetFilePointer fh, 0, 0, FILE_BEGIN
ReadFile fh, TagHeader, Len(TagHeader), 0, ByVal 0
If Data2String(VarPtr(TagHeader.Identifier(0)), 3, ENC_ISO) = "ID3" Then
    outInfo.ID3v2Version = TagHeader.Version(0)
End If

'File size.
outInfo.FileSize = fsize

CloseHandle fh
End Function
Private Function GetTagSize(ByVal fh As Long, ByVal fp As Long) As Long
Dim TagHeader As v2TagHeader

'Search for an ID3v2 tag.
SetFilePointer fh, fp, 0, FILE_BEGIN
ReadFile fh, TagHeader, Len(TagHeader), 0, ByVal 0
If Data2String(VarPtr(TagHeader.Identifier(0)), 3, ENC_ISO) = "ID3" Then
    'The size stored in the header excludes itself, and excludes the footer (if present).
    GetTagSize = Data2Long(VarPtr(TagHeader.Size(0)), True) + Len(TagHeader)

    'v 2.4 (or later?) flags: %abcd0000 abc = ignored, d = footer present
    If TagHeader.Version(0) >= 4 Then
        If TagHeader.flags And &H10& Then
            'Add the size of the footer (which is the same size than the header) to the existing size.
            GetTagSize = GetTagSize + Len(TagHeader)
        End If
    End If
End If
End Function

Private Sub IgnoreReadOnly(ByVal FileName As String)
Dim rdAttr As Long

rdAttr = GetFileAttributes(FileName)
If rdAttr = -1 Then Exit Sub
If rdAttr And FILE_ATTRIBUTE_READONLY Then
    'Set this bit to 0. Use xor for that.
    rdAttr = rdAttr Xor FILE_ATTRIBUTE_READONLY
    SetFileAttributes FileName, rdAttr
End If
End Sub

Public Sub ResetTag(ByRef Tag As ID3Tag)
With Tag
    .Album = ""
    .Artist = ""
    .Comment = ""
    .Composer = ""
    .Copyright = ""
    .EncodedBy = ""
    .Genre = ""
    .LinkTo = ""
    .OrigArtist = ""
    .SongYear = ""
    .Title = ""
    .TrackNr = ""
End With
End Sub
Private Sub Class_Initialize()
ID3_GENRE_COUNT = 148
End Sub


