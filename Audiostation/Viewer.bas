Attribute VB_Name = "Viewer"
Public Const MAX_PATH = 260
   Public Const SHGFI_DISPLAYNAME = &H200
   Public Const SHGFI_EXETYPE = &H2000
   Public Const SHGFI_SYSICONINDEX = &H4000 'system icon index
   Public Const SHGFI_LARGEICON = &H0 'large icon
   Public Const SHGFI_SMALLICON = &H1 'small icon
   Public Const ILD_TRANSPARENT = &H1 'display transparent
   Public Const SHGFI_SHELLICONSIZE = &H4
   Public Const SHGFI_TYPENAME = &H400
   Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
   Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
   Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE


   Public Type SHFILEINFO
       hIcon As Long
       iIcon As Long
       dwAttributes As Long
       szDisplayName As String * MAX_PATH
       szTypeName As String * 80
       End Type


   Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
       (ByVal pszPath As String, _
       ByVal dwFileAttributes As Long, _
       psfi As SHFILEINFO, _
       ByVal cbSizeFileInfo As Long, _
       ByVal uFlags As Long) As Long


   Public Declare Function ImageList_Draw Lib "comctl32.dll" _
       (ByVal himl&, ByVal i&, ByVal hDCDest&, _
       ByVal x&, ByVal Y&, ByVal flags&) As Long
       Public shinfo As SHFILEINFO


