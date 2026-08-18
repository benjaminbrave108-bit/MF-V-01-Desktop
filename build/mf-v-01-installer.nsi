Unicode true
RequestExecutionLevel user

!include "MUI2.nsh"

!define APP_NAME "MF-V-01"
!define APP_VERSION "1.0.7"
!define APP_EXE "MF-V-01.exe"
!define APP_DIR "$LOCALAPPDATA\Programs\MF-V-01"

Name "${APP_NAME} ${APP_VERSION}"
OutFile "..\release\MF-V-01-Setup-1.0.7.exe"
InstallDir "${APP_DIR}"
InstallDirRegKey HKCU "Software\KB Group\MF-V-01" "InstallLocation"
ShowInstDetails show
ShowUninstDetails show

!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\${APP_EXE}"
!define MUI_FINISHPAGE_RUN_TEXT "Open MF-V-01 / MF-V-01'i aç / MF-V-01 veke"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "Turkish"

Section "MF-V-01" SEC_APP
  SetShellVarContext current
  SetOutPath "$INSTDIR"
  File /r "..\release\win-unpacked\*.*"

  WriteUninstaller "$INSTDIR\Uninstall MF-V-01.exe"
  CreateDirectory "$SMPROGRAMS\MF-V-01"
  CreateShortcut "$SMPROGRAMS\MF-V-01\MF-V-01.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}" 0
  CreateShortcut "$DESKTOP\MF-V-01.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}" 0

  WriteRegStr HKCU "Software\KB Group\MF-V-01" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MF-V-01" "DisplayName" "MF-V-01 ${APP_VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MF-V-01" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MF-V-01" "Publisher" "KB Group"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MF-V-01" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MF-V-01" "DisplayIcon" "$INSTDIR\${APP_EXE}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MF-V-01" "UninstallString" '$\"$INSTDIR\Uninstall MF-V-01.exe$\"'
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MF-V-01" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MF-V-01" "NoRepair" 1
SectionEnd

Section "Uninstall"
  SetShellVarContext current
  Delete "$DESKTOP\MF-V-01.lnk"
  Delete "$SMPROGRAMS\MF-V-01\MF-V-01.lnk"
  RMDir "$SMPROGRAMS\MF-V-01"

  Delete "$INSTDIR\Uninstall MF-V-01.exe"
  RMDir /r "$INSTDIR"

  DeleteRegKey HKCU "Software\KB Group\MF-V-01"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MF-V-01"
SectionEnd
