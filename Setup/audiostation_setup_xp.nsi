!define PRODUCT_NAME "Audiostation"
!define PRODUCT_VERSION "2.4.0"
!define PRODUCT_PUBLISHER "Sibra-Soft"
!define PRODUCT_WEB_SITE "https://www.audiostation.org"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\Audiostation.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; MUI 1.67 compatible ------
!include "MUI.nsh"

; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON ".\program.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Language Selection Dialog Settings
!define MUI_LANGDLL_REGISTRY_ROOT "${PRODUCT_UNINST_ROOT_KEY}"
!define MUI_LANGDLL_REGISTRY_KEY "${PRODUCT_UNINST_KEY}"
!define MUI_LANGDLL_REGISTRY_VALUENAME "NSIS:Language"

!define MUI_HEADERIMAGE_BITMAP ".\header.bmp"
!define MUI_UI_HEADERIMAGE_RIGHT ".\header.bmp"
!define MUI_WELCOMEFINISHPAGE_BITMAP ".\wizard.bmp"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_RUN "$INSTDIR\Audiostation.exe"
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "Dutch"
!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "German"

RequestExecutionLevel admin

; MUI end ------

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "Ministation_windows_setup.exe"
InstallDir "$PROGRAMFILES\Sibra-Soft\Ministation"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd

Section "Audiostation" SEC01
  SetOutPath "$INSTDIR"

  SetOverwrite ifnewer
  SetOverwrite try
  
  CreateDirectory "$SMPROGRAMS\Ministation"

  File ".\build\Audiostation.exe"
  File ".\deps\AdioLibrary.ocx"
  File ".\deps\bass.dll"
  File ".\deps\basscd.dll"
  File ".\deps\bassflac.dll"
  File ".\deps\bassmix.dll"
  File ".\deps\basswasapi.dll"
  File ".\deps\comdlg32.ocx"
  File ".\deps\d3DLine.ocx"
  File ".\deps\isAnalogLibrary.ocx"
  File ".\deps\isDigitalLibrary.ocx"
  File ".\deps\MBPrgBar.ocx"
  File ".\deps\mscomctl.ocx"

  RegDLL "$INSTDIR\AdioLibrary.ocx"
  RegDLL "$INSTDIR\bass.dll"
  RegDLL "$INSTDIR\basscd.dll"
  RegDLL "$INSTDIR\bassflac.dll"
  RegDLL "$INSTDIR\bassmix.dll"
  RegDLL "$INSTDIR\basswasapi.dll"
  RegDLL "$INSTDIR\comdlg32.ocx"
  RegDLL "$INSTDIR\d3DLine.ocx"
  RegDLL "$INSTDIR\isAnalogLibrary.ocx"
  RegDLL "$INSTDIR\isDigitalLibrary.ocx"
  RegDLL "$INSTDIR\MBPrgBar.ocx"
  RegDLL "$INSTDIR\mscomctl.ocx"
SectionEnd

Section "Start Menu Shortcuts" SEC02
  WriteIniStr "$INSTDIR\${PRODUCT_NAME}.url" "InternetShortcut" "URL" "${PRODUCT_WEB_SITE}"
  CreateShortCut "$SMPROGRAMS\Audiostation\Website.lnk" "$INSTDIR\${PRODUCT_NAME}.url"
  CreateShortCut "$SMPROGRAMS\Audiostation\Uninstall.lnk" "$INSTDIR\uninst.exe"
  CreateShortCut "$SMPROGRAMS\Audiostation\Ministation.lnk" "$INSTDIR\Audiostation.exe"
SectionEnd

Section "Desktop Shortcut" SEC03
  CreateShortCut "$DESKTOP\Ministation.lnk" "$INSTDIR\Audiostation.exe"
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\Audiostation.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\Audiostation.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd

Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) was successfully removed from your computer."
FunctionEnd

Function un.onInit
!insertmacro MUI_UNGETLANGUAGE
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "Are you sure you want to completely remove $(^Name) and all of its components?" IDYES +2
  Abort
FunctionEnd

Section Uninstall
  Delete "$INSTDIR\${PRODUCT_NAME}.url"
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\mscomctl.ocx"
  Delete "$INSTDIR\MBPrgBar.ocx"
  Delete "$INSTDIR\isDigitalLibrary.ocx"
  Delete "$INSTDIR\isAnalogLibrary.ocx"
  Delete "$INSTDIR\d3DLine.ocx"
  Delete "$INSTDIR\comdlg32.ocx"
  Delete "$INSTDIR\basswasapi.dll"
  Delete "$INSTDIR\bassmix.dll"
  Delete "$INSTDIR\bassflac.dll"
  Delete "$INSTDIR\basscd.dll"
  Delete "$INSTDIR\bass.dll"
  Delete "$INSTDIR\AdioLibrary.ocx"
  Delete "$INSTDIR\Ministation.exe"

  Delete "$SMPROGRAMS\Audiostation\Uninstall.lnk"
  Delete "$SMPROGRAMS\Audiostation\Website.lnk"
  Delete "$DESKTOP\Audiostation.lnk"
  Delete "$SMPROGRAMS\Audiostation\Audiostation.lnk"

  RMDir "$SMPROGRAMS\Audiostation"
  RMDir "$INSTDIR"

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  SetAutoClose true
SectionEnd