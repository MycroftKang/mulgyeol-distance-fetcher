!define PRODUCT_PUBLISHER "Mulgyeol Labs"
!define PRODUCT_WEB_SITE "https://github.com/MycroftKang/mulgyeol-distance-fetcher/releases"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\${EXE_NAME}"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKCU"

RequestExecutionLevel user
Unicode true

; MUI 1.67 compatible ------
!include "MUI.nsh"
!include "FileFunc.nsh"

; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON "..\resources\package\installer_icon.ico"
!define MUI_UNICON "..\resources\package\installer_icon.ico"

;!define MUI_PAGE_HEADER_TEXT "Mulgyeol Software Update"
;!define MUI_PAGE_HEADER_SUBTEXT "MDFE 3.1.3"

BrandingText "Mulgyeol Labs"

; Welcome page
!define MUI_WELCOMEFINISHPAGE_BITMAP "..\resources\package\welcome.bmp"
!define MUI_WELCOMEPAGE_TITLE "${PRODUCT_NAME} ${PRODUCT_VERSION} 설치를 시작합니다."
; !define MUI_WELCOMEPAGE_TEXT "3.2.0 새로운 기능\n\n- 최근 항목 불러오기 기능 추가 및 업데이터 개선 등\n\n3.1.3 새로운 기능\n\n- 네트워크 오류 감지 알고리즘이 개선"
!insertmacro MUI_PAGE_WELCOME

; License page
!insertmacro MUI_PAGE_LICENSE "..\LICENSE"

; Directory page
!insertmacro MUI_PAGE_DIRECTORY

; Instfiles page
!insertmacro MUI_PAGE_INSTFILES

; Finish page
!define MUI_FINISHPAGE_RUN
!define MUI_FINISHPAGE_RUN_CHECKED
!define MUI_FINISHPAGE_RUN_FUNCTION "RunMDF"
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "Korean"

; MUI end ------

; Caption "Mulgyeol Software Update - MDFE ${PRODUCT_VERSION}"
Name "${PRODUCT_NAME}"
OutFile "MDFSetup-stable.exe"
InstallDir "$LOCALAPPDATA\Programs\${PRODUCT_NAME}"

Function RunMDF
  SetOutPath "$INSTDIR\app"
  Exec "$INSTDIR\app\${EXE_NAME}"
FunctionEnd

Function .onInit
StrCpy $INSTDIR "$LOCALAPPDATA\Programs\${PRODUCT_NAME}"  
iffileexists "$ProgramFiles\${PRODUCT_NAME}\uninst.exe" YES NO
  YES:
  MessageBox MB_OK "이전 버전의 ${PRODUCT_NAME}를 제거한 후 다시 시도하십시오." /SD IDOK
  IfSilent +3
  ExecShell "runas" "$ProgramFiles\${PRODUCT_NAME}\uninst.exe"
  Abort
  ExecShellWait "runas" "$ProgramFiles\${PRODUCT_NAME}\uninst.exe /S"
  NO:
FunctionEnd

InstallDirRegKey HKCU "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

Section "app" SEC01
  ; nsExec::Exec 'taskkill /f /im "Mulgyeol Software Update.exe"'
  SetOutPath "$INSTDIR"
  File "msu.exe"
  IfSilent +1 +2
  Exec "$INSTDIR\msu.exe /start"
  File "..\product.json"
  SetOutPath "$INSTDIR\app"
  File "MDF.VisualElementsManifest.xml"
  File /nonfatal /a /r "..\build\app\*"
  CreateDirectory "$SMPROGRAMS\${PRODUCT_NAME}"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\${PRODUCT_NAME}.lnk" "$INSTDIR\app\${EXE_NAME}"
  CreateShortCut "$DESKTOP\${PRODUCT_NAME}.lnk" "$INSTDIR\app\${EXE_NAME}"
  SetOutPath "$INSTDIR\update"
  File /nonfatal /a /r "..\build\update\*"
  SetOutPath "$INSTDIR\resources\app"
  File /nonfatal /a /r "..\resources\app\*"
  SetOutPath "$INSTDIR\app\visual"
  File /nonfatal /a /r "..\resources\visual\*"
  SetOutPath "$INSTDIR\info"
  File /nonfatal /a /r "info\*"
  SetOutPath "$INSTDIR\data"
  SetOverwrite off
  File /nonfatal /a /r "data\*"
  SetOverwrite on
SectionEnd

Section -AdditionalIcons
  SetOutPath $INSTDIR
  WriteIniStr "$INSTDIR\${PRODUCT_NAME}.url" "InternetShortcut" "URL" "${PRODUCT_WEB_SITE}"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\Website.lnk" "$INSTDIR\${PRODUCT_NAME}.url"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\Uninstall.lnk" "$INSTDIR\uninst.exe"
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKCU "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\app\${EXE_NAME}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\app\${EXE_NAME}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd

Function .onInstSuccess
  IfSilent +1 +2
  nsExec::Exec 'taskkill /f /im "msu.exe"'
  ${GetParameters} $1
  ClearErrors
  ${GetOptions} $1 '/autorun' $R0
  IfErrors +3 0
  SetOutPath "$INSTDIR\app"
  Exec "$INSTDIR\app\${EXE_NAME}"
FunctionEnd

Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name)는(은) 완전히 제거되었습니다." /SD IDOK
FunctionEnd

Function un.onInit
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "$(^Name)을(를) 제거하시겠습니까?" /SD IDYES IDYES +2
  Abort
FunctionEnd

Section Uninstall
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\Uninstall.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\Website.lnk"
  Delete "$DESKTOP\${PRODUCT_NAME}.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\${PRODUCT_NAME}.lnk"

  RMDir "$SMPROGRAMS\${PRODUCT_NAME}"
  RMDir /r /REBOOTOK $INSTDIR

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKCU "${PRODUCT_DIR_REGKEY}"
  SetAutoClose true
SectionEnd