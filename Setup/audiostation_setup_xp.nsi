!define PRODUCT_NAME "Audiostation"
!define PRODUCT_VERSION "2.4.1"
!define PRODUCT_PUBLISHER "Sibra-Soft"
!define PRODUCT_WEB_SITE "https://www.audiostation.org"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\Audiostation.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

!include "WinVer.nsh"
!include "MUI.nsh"
!include ".\fileassoc.nsh"
!include ".\lang.nsh"

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
!define MUI_COMPONENTSPAGE_NODESC

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

Name "${PRODUCT_NAME}"
OutFile "Audiostation_windows_xp_setup.exe"
InstallDir "$PROGRAMFILES\Sibra-Soft\Audiostation"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
  ${If} ${AtLeastWinVista}
    MessageBox MB_ICONSTOP|MB_OK "This installer only supports Windows XP.$\r$\nPlease use a newer version of this software."
    Abort
  ${EndIf}
FunctionEnd

Section "!Audiostation" SEC_main
  SetOverwrite ifnewer
  SetOverwrite try

  CreateDirectory "$SMPROGRAMS\Audiostation"

  SetOutPath "$INSTDIR\languages"
  File ".\languages\dutch.lng"
  File ".\languages\english.lng"
  File ".\languages\german.lng"

  SetOutPath "$INSTDIR"
  File ".\build\Audiostation.exe"

  File ".\publish\win32.dll"
  File ".\publish\settings.ini"
  File ".\publish\recorder.exe"
  File ".\publish\streams.db"

  File ".\deps\AdioLibrary.ocx"
  File ".\deps\bass.dll"
  File ".\deps\basscd.dll"
  File ".\deps\bassflac.dll"
  File ".\deps\bassmix.dll"
  File ".\deps\basswasapi.dll"
  File ".\deps\bass_aac.dll"
  File ".\deps\COMCT232.OCX"
  File ".\deps\Comctl32.ocx"
  File ".\deps\COMDLG32.OCX"
  File ".\deps\d3DLine.ocx"
  File ".\deps\DataInter.ocx"
  File ".\deps\DigitBox.ocx"
  File ".\deps\DirDlg.ocx"
  File ".\deps\isAnalogLibrary.ocx"
  File ".\deps\isDigitalLibrary.ocx"
  File ".\deps\LaVolpeAlphaImg2.ocx"
  File ".\deps\MBPrgBar.ocx"
  File ".\deps\midifl2k.ocx"
  File ".\deps\midifl32.ocx"
  File ".\deps\midiio2k.ocx"
  File ".\deps\midiio32.ocx"
  File ".\deps\mscomctl.ocx"
  File ".\deps\MSVCRT40.DLL"
  File ".\deps\TABCTL32.OCX"
  File ".\deps\wshom.ocx"

  RegDLL "$INSTDIR\AdioLibrary.ocx"
  RegDLL "$INSTDIR\COMCT232.OCX"
  RegDLL "$INSTDIR\Comctl32.ocx"
  RegDLL "$INSTDIR\COMDLG32.OCX"
  RegDLL "$INSTDIR\d3DLine.ocx"
  RegDLL "$INSTDIR\DataInter.ocx"
  RegDLL "$INSTDIR\DigitBox.ocx"
  RegDLL "$INSTDIR\DirDlg.ocx"
  RegDLL "$INSTDIR\isAnalogLibrary.ocx"
  RegDLL "$INSTDIR\isDigitalLibrary.ocx"
  RegDLL "$INSTDIR\LaVolpeAlphaImg2.ocx"
  RegDLL "$INSTDIR\MBPrgBar.ocx"
  RegDLL "$INSTDIR\midifl2k.ocx"
  RegDLL "$INSTDIR\midifl32.ocx"
  RegDLL "$INSTDIR\midiio2k.ocx"
  RegDLL "$INSTDIR\midiio32.ocx"
  RegDLL "$INSTDIR\mscomctl.ocx"
  RegDLL "$INSTDIR\TABCTL32.OCX"
  RegDLL "$INSTDIR\wshom.ocx"
SectionEnd

Section "Start Menu Shortcuts" SEC02
  WriteIniStr "$INSTDIR\${PRODUCT_NAME}.url" "InternetShortcut" "URL" "${PRODUCT_WEB_SITE}"
  CreateShortCut "$SMPROGRAMS\Audiostation\Website.lnk" "$INSTDIR\${PRODUCT_NAME}.url"
  CreateShortCut "$SMPROGRAMS\Audiostation\Uninstall.lnk" "$INSTDIR\uninst.exe"
  CreateShortCut "$SMPROGRAMS\Audiostation\Audiostation.lnk" "$INSTDIR\Audiostation.exe"
SectionEnd

Section "Desktop Shortcut" SEC03
  CreateShortCut "$DESKTOP\Audiostation.lnk" "$INSTDIR\Audiostation.exe"
SectionEnd

Section "Sample Files" SEC04
  SetOverwrite ifnewer
  SetOverwrite try

  CreateDirectory "$INSTDIR\temp"

  SetOutPath "$INSTDIR\sampels"
  File ".\publish\samples\DEMO16.WAV"
  File ".\publish\samples\DEMOWSS.WAV"
  SetOutPath "$INSTDIR\sampels\MIDFILES"
  File ".\publish\samples\MIDFILES\8NOTERWD.MID"
  File ".\publish\samples\MIDFILES\BALLADE.MID"
  File ".\publish\samples\MIDFILES\BARBER.MID"
  File ".\publish\samples\MIDFILES\BLUEDANU.MID"
  File ".\publish\samples\MIDFILES\DEMO1.MID"
  File ".\publish\samples\MIDFILES\DRUM1.MID"
  File ".\publish\samples\MIDFILES\ENTERTN.MID"
  File ".\publish\samples\MIDFILES\GROOVE.MID"
  File ".\publish\samples\MIDFILES\HIP2.MID"
  File ".\publish\samples\MIDFILES\ITSHOP.MID"
  File ".\publish\samples\MIDFILES\JAZZ.MID"
  File ".\publish\samples\MIDFILES\KOOLTHIN.MID"
  File ".\publish\samples\MIDFILES\LISZT2.MID"
  File ".\publish\samples\MIDFILES\MARCH.MID"
  File ".\publish\samples\MIDFILES\MINUET.MID"
  File ".\publish\samples\MIDFILES\REGGAE.MID"
  File ".\publish\samples\MIDFILES\SERENITY.MID"
  File ".\publish\samples\MIDFILES\SNDO.MID"
  File ".\publish\samples\MIDFILES\SONTINA3.MID"
  File ".\publish\samples\MIDFILES\STACCATO.MID"
  File ".\publish\samples\MIDFILES\TURKISH.MID"
  File ".\publish\samples\MIDFILES\TUTOR1.MID"
  File ".\publish\samples\MIDFILES\WALTZFLR.MID"
  File ".\publish\samples\MIDFILES\WILLTELL.MID"
  File ".\publish\samples\MIDFILES\WIL_TELL.MID"
  File ".\publish\samples\MIDFILES\WOODEN.MID"
  SetOutPath "$INSTDIR\sampels\MUSFILES"
  File ".\publish\samples\MUSFILES\Mountain.mus"
  File ".\publish\samples\MUSFILES\SuperMario 2.mus"
  File ".\publish\samples\MUSFILES\SuperMario.mus"
  File ".\publish\samples\MUSFILES\Tetris.mus"
  File ".\publish\samples\MUSFILES\William_Tell.mus"
  SetOutPath "$INSTDIR\sampels\WAVFILES"
  File ".\publish\samples\WAVFILES\APPLAUS.WAV"
  File ".\publish\samples\WAVFILES\BABYCRY.WAV"
  File ".\publish\samples\WAVFILES\BAGPIPE.WAV"
  File ".\publish\samples\WAVFILES\BALLGAME.WAV"
  File ".\publish\samples\WAVFILES\BIGBASS.WAV"
  File ".\publish\samples\WAVFILES\BROOK.WAV"
  File ".\publish\samples\WAVFILES\BUBBLES.WAV"
  File ".\publish\samples\WAVFILES\CAROSEL.WAV"
  File ".\publish\samples\WAVFILES\CHASE.WAV"
  File ".\publish\samples\WAVFILES\CHOOCHOO.WAV"
  File ".\publish\samples\WAVFILES\CHORUS.WAV"
  File ".\publish\samples\WAVFILES\CRASH.WAV"
  File ".\publish\samples\WAVFILES\DOORSLAM.WAV"
  File ".\publish\samples\WAVFILES\DOORSQUK.WAV"
  File ".\publish\samples\WAVFILES\DRUMCORP.WAV"
  File ".\publish\samples\WAVFILES\HELICOPT.WAV"
  File ".\publish\samples\WAVFILES\JUNGLDRM.WAV"
  File ".\publish\samples\WAVFILES\REVCYMB.WAV"
  File ".\publish\samples\WAVFILES\SCARY.WAV"
  File ".\publish\samples\WAVFILES\SCREAM.WAV"
  File ".\publish\samples\WAVFILES\STORMY.WAV"
  File ".\publish\samples\WAVFILES\TIMPS.WAV"
  File ".\publish\samples\WAVFILES\TRUMPETX.WAV"
SectionEnd

Section "VirtualMIDISynth" SEC05
  SetOutPath "$INSTDIR"
  SetOverwrite on
  File ".\publish\VirtualMIDISynth-2540.exe"
SectionEnd

Section "File Associations" SEC06
  !insertmacro APP_ASSOCIATE "mp3" "audiostation.mp3" "MP3" "$INSTDIR\win32.dll,10" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "wav" "audiostation.wav" "WAV" "$INSTDIR\win32.dll,14" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "mid" "audiostation.mid" "MID" "$INSTDIR\win32.dll,8" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "wma" "audiostation.wma" "WMA" "$INSTDIR\win32.dll,18" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "kar" "audiostation.kar" "KAR" "$INSTDIR\win32.dll,5" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "mp2" "audiostation.mp2" "MP2" "$INSTDIR\win32.dll,9" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "aac" "audiostation.aac" "AAC" "$INSTDIR\win32.dll,0" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "snd" "audiostation.snd" "SND" "$INSTDIR\win32.dll,16" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "au" "audiostation.au" "AU" "$INSTDIR\win32.dll,4" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "rmi" "audiostation.rmi" "RMI" "$INSTDIR\win32.dll,14" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "m4a" "audiostation.m4a" "M4A" "$INSTDIR\win32.dll,9" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "cda" "audiostation.cda" "CDA" "$INSTDIR\win32.dll,4" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "ra" "audiostation.ra" "RA" "$INSTDIR\win32.dll,0" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "mus" "audiostation.mus" "MUS" "$INSTDIR\win32.dll,11" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "sid" "audiostation.sid" "SID" "$INSTDIR\win32.dll,13" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "apl" "audiostation.apl" "Audiostation Playlist" "$INSTDIR\win32.dll,1" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "pls" "audiostation.pls" "PLS" "$INSTDIR\win32.dll,12" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "m3u" "audiostation.m3u" "M3U" "$INSTDIR\win32.dll,6" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
  !insertmacro APP_ASSOCIATE "wpl" "audiostation.wpl" "Windows Media Player Playlist" "$INSTDIR\win32.dll,19" "Open with Audiostation" "$INSTDIR\audiostation.exe $\"%1$\""
SectionEnd

Section -Post
  !insertmacro LanguageCodeToText $LANGUAGE $0

  WriteINIStr "$INSTDIR\settings.ini" "main" "Langauge" "$0"

  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\Audiostation.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\Audiostation.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"

  Exec "$INSTDIR\VirtualMIDISynth-2540.exe"
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

  Delete "$INSTDIR\Audiostation.exe"

  Delete "$INSTDIR\win32.dll"
  Delete "$INSTDIR\settings.ini"
  Delete "$INSTDIR\recorder.exe"
  Delete "$INSTDIR\streams.db"

  Delete "$INSTDIR\languages\dutch.lng"
  Delete "$INSTDIR\languages\english.lng"
  Delete "$INSTDIR\languages\german.lng"

  Delete "$INSTDIR\VirtualMIDISynth-21370.exe"

  Delete "$INSTDIR\bass.dll"
  Delete "$INSTDIR\bass_aac.dll"
  Delete "$INSTDIR\basscd.dll"
  Delete "$INSTDIR\bassflac.dll"
  Delete "$INSTDIR\bassmix.dll"
  Delete "$INSTDIR\basswasapi.dll"
  Delete "$INSTDIR\AdioLibrary.ocx"
  Delete "$INSTDIR\COMCT232.OCX"
  Delete "$INSTDIR\Comctl32.ocx"
  Delete "$INSTDIR\COMDLG32.OCX"
  Delete "$INSTDIR\d3DLine.ocx"
  Delete "$INSTDIR\DataInter.ocx"
  Delete "$INSTDIR\DigitBox.ocx"
  Delete "$INSTDIR\DirDlg.ocx"
  Delete "$INSTDIR\isAnalogLibrary.ocx"
  Delete "$INSTDIR\isDigitalLibrary.ocx"
  Delete "$INSTDIR\LaVolpeAlphaImg2.ocx"
  Delete "$INSTDIR\MBPrgBar.ocx"
  Delete "$INSTDIR\midifl2k.ocx"
  Delete "$INSTDIR\midifl32.ocx"
  Delete "$INSTDIR\midiio2k.ocx"
  Delete "$INSTDIR\midiio32.ocx"
  Delete "$INSTDIR\mscomctl.ocx"
  Delete "$INSTDIR\TABCTL32.OCX"
  Delete "$INSTDIR\wshom.ocx"

  Delete "$SMPROGRAMS\Audiostation\Uninstall.lnk"
  Delete "$SMPROGRAMS\Audiostation\Website.lnk"
  Delete "$DESKTOP\Audiostation.lnk"
  Delete "$SMPROGRAMS\Audiostation\Audiostation.lnk"

  RMDir "$SMPROGRAMS\Audiostation"
  RMDir "$INSTDIR\languages"
  RMDir "$INSTDIR\sampels"
  RMDir "$INSTDIR"

  !insertmacro APP_UNASSOCIATE "mp3" "audiostation.mp3"
  !insertmacro APP_UNASSOCIATE "wav" "audiostation.wav"
  !insertmacro APP_UNASSOCIATE "mid" "audiostation.mid"
  !insertmacro APP_UNASSOCIATE "wma" "audiostation.wma"
  !insertmacro APP_UNASSOCIATE "kar" "audiostation.kar"
  !insertmacro APP_UNASSOCIATE "mp2" "audiostation.mp2"
  !insertmacro APP_UNASSOCIATE "aac" "audiostation.aac"
  !insertmacro APP_UNASSOCIATE "snd" "audiostation.snd"
  !insertmacro APP_UNASSOCIATE "au" "audiostation.au"
  !insertmacro APP_UNASSOCIATE "rmi" "audiostation.rmi"
  !insertmacro APP_UNASSOCIATE "m4a" "audiostation.m4a"
  !insertmacro APP_UNASSOCIATE "cda" "audiostation.cda"
  !insertmacro APP_UNASSOCIATE "ra" "audiostation.ra"
  !insertmacro APP_UNASSOCIATE "mus" "audiostation.mus"
  !insertmacro APP_UNASSOCIATE "sid" "audiostation.sid"
  !insertmacro APP_UNASSOCIATE "apl" "audiostation.apl"
  !insertmacro APP_UNASSOCIATE "pls" "audiostation.pls"
  !insertmacro APP_UNASSOCIATE "m3u" "audiostation.m3u"
  !insertmacro APP_UNASSOCIATE "wpl" "audiostation.wpl"

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"

  SetAutoClose true
SectionEnd