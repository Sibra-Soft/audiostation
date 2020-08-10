VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA1Hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'SHA1Hash
'
'Perform CryptoAPI SHA1 hash of supplied file, returning hash
'as string of 40 hex digits.

'----- Private Consts -----

Private Const ALG_TYPE_ANY          As Long = 0
Private Const ALG_CLASS_HASH        As Long = 32768
Private Const ALG_SID_SHA1          As Long = 4
Private Const CALG_SHA1             As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1

Private Const PROV_RSA_FULL         As Long = 1
Private Const CRYPT_VERIFYCONTEXT   As Long = &HF0000000
Private Const MS_DEFAULT_PROVIDER   As String = _
    "Microsoft Base Cryptographic Provider v1.0"

Private Const HP_HASHVAL            As Long = 2
Private Const HP_HASHSIZE           As Long = 4
                  
'----- Private Defines -----

Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextA" ( _
    ByRef phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long 'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.

Private Declare Function CryptCreateHash Lib "advapi32" ( _
    ByVal hProv As Long, _
    ByVal algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phHash As Long) As Long 'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.
    
Private Declare Function CryptDestroyHash Lib "advapi32" ( _
    ByVal hHash As Long) As Long 'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.

Private Declare Function CryptGetHashParam Lib "advapi32" ( _
    ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptHashData Lib "advapi32" ( _
    ByVal hHash As Long, _
    ByRef pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptReleaseContext Lib "advapi32" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long 'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.

'----- Private Data -----

Private m_hHash As Long 'Hash object handle.
Private m_hProvider As Long 'Cryptographic Service Provider handle.

'----- Private Methods -----

Private Sub HashBlock(ByRef Block() As Byte)
Attribute HashBlock.VB_Description = "Hash a block of data"
    If CryptHashData(m_hHash, _
                     Block(LBound(Block)), _
                     UBound(Block) - LBound(Block) + 1, _
                     0&) = 0 Then
        Err.Raise vbObjectError Or &HC312&, _
                  "SHA1Hash.HashFile", _
                  "Failed to hash data block, system error " _
                & CStr(Err.LastDllError)
    End If
End Sub

Private Function HashValue() As String
Attribute HashValue.VB_Description = "Retrieve hash value and terminate hashing sequence"
    Dim lngDataLen As Long
    Dim lngHashSize As Long
    Dim bytHashValue() As Byte
    
    lngDataLen = 4 '4 bytes for Long length.
    If CryptGetHashParam(m_hHash, HP_HASHSIZE, lngHashSize, lngDataLen, 0&) = 0 Then
        Err.Raise vbObjectError Or &HC322&, _
                  "SHA1Hash.HashFile", _
                  "Failed to obtain hash value length, system error " _
                & CStr(Err.LastDllError)
    Else
        lngDataLen = lngHashSize
        ReDim bytHashValue(lngDataLen - 1)
        
        If CryptGetHashParam(m_hHash, HP_HASHVAL, bytHashValue(0), lngDataLen, 0&) = 0 Then
            Err.Raise vbObjectError Or &HC324&, _
                      "SHA1Hash.HashFile", _
                      "Failed to obtain hash value, system error " _
                    & CStr(Err.LastDllError)
        Else
            Dim intByte As Integer
            
            For intByte = 0 To lngDataLen - 1
                HashValue = HashValue & Right$("0" & Hex$(bytHashValue(intByte)), 2)
            Next
            
            CryptDestroyHash m_hHash
        End If
    End If
End Function

Private Sub NewHash()
Attribute NewHash.VB_Description = "Initialize a new hashing sequence"
    If CryptCreateHash(m_hProvider, CALG_SHA1, 0&, 0&, m_hHash) = 0 Then
        Err.Raise vbObjectError Or &HC332&, _
                  "SHA1Hash.HashFile", _
                  "Failed to create CryptoAPI Hash object, system error " _
                & CStr(Err.LastDllError)
    End If
End Sub

'----- Public Methods -----

Public Function HashFile(ByVal FileName As String) As String
Attribute HashFile.VB_Description = "Return hash of file's contents"
    Const CHUNK As Long = 16384
    Dim intFile As Integer
    Dim lngWholeChunks As Long
    Dim intRemainder As Integer
    Dim lngChunk As Long
    Dim bytBlock() As Byte
    
    
    If Len(Dir$(FileName, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)) > 0 Then
        intFile = FreeFile(0)
        Open FileName For Binary Access Read As #intFile
        lngWholeChunks = LOF(intFile) \ CHUNK
        intRemainder = LOF(intFile) - (CHUNK * lngWholeChunks)
        NewHash
        ReDim bytBlock(CHUNK - 1)
        For lngChunk = 1 To lngWholeChunks
            Get #intFile, , bytBlock
            HashBlock bytBlock
        Next
        If intRemainder > 0 Then
            ReDim bytBlock(intRemainder - 1)
            Get #intFile, , bytBlock
            HashBlock bytBlock
        End If
        Close #intFile
        HashFile = HashValue()
    Else
        Err.Raise vbObjectError Or &HC342&, _
                  "SHA1Hash.HashFile", _
                  "File doesn't exist"
    End If
End Function

Public Function HashBytes(ByRef Bytes() As Byte) As String
    NewHash
    HashBlock Bytes
    HashBytes = HashValue()
End Function

'----- Class Event Handlers -----

Private Sub Class_Initialize()
    If CryptAcquireContext(m_hProvider, _
                           vbNullString, _
                           MS_DEFAULT_PROVIDER, _
                           PROV_RSA_FULL, _
                           CRYPT_VERIFYCONTEXT) = 0 Then
        Err.Raise vbObjectError Or &HC352&, _
                  "SHA1Hash.Class_Initialize", _
                  "Failed to obtain access to CryptoAPI, system error " _
                & CStr(Err.LastDllError)
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next 'All exceptions must be processed here.
    CryptDestroyHash m_hHash
    CryptReleaseContext m_hProvider, 0&
End Sub
