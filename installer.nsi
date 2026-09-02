Unicode true
Name "조사연구 도우미"
OutFile "dist\ResearchReportHelperSetup.exe"
InstallDir "$LOCALAPPDATA\Programs\ResearchReportHelper"
RequestExecutionLevel user

; 버전의 단일 원본은 updater.py 의 APP_VERSION 이다. build_installer.ps1 이 거기서 읽어
; /DAPP_VERSION=... 으로 넘긴다. 직접 makensis 를 돌릴 때를 위한 폴백만 아래에 둔다.
!ifndef APP_VERSION
  !define APP_VERSION "2.0.6"
!endif

!define APP_EXE "ResearchReportHelper.exe"
!define OLD_APP_EXE "ResearchReportAutomation.exe"
!define APP_NAME "조사연구 도우미"
!define OLD_APP_NAME "Research Report Automation"
!define COMPANY_NAME "Jeonbuk Bank AI Innovation Department"
!define PUBLISHER_NAME "전북은행 AI혁신부"
!define APP_DATA_DIR "$LOCALAPPDATA\ResearchReportAutomation"

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

Function .onInit
  nsExec::ExecToLog 'taskkill /F /T /IM "${APP_EXE}"'
  nsExec::ExecToLog 'taskkill /F /T /IM "${OLD_APP_EXE}"'
FunctionEnd

Function un.onInit
  nsExec::ExecToLog 'taskkill /F /T /IM "${APP_EXE}"'
  nsExec::ExecToLog 'taskkill /F /T /IM "${OLD_APP_EXE}"'
FunctionEnd

Section "Install"
  nsExec::ExecToLog 'taskkill /F /T /IM "${APP_EXE}"'
  nsExec::ExecToLog 'taskkill /F /T /IM "${OLD_APP_EXE}"'
  Sleep 1000
  SetOutPath "$INSTDIR"
  SetOverwrite on
  File "dist\${APP_EXE}"

  WriteUninstaller "$INSTDIR\Uninstall.exe"
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\삭제.lnk" "$INSTDIR\Uninstall.exe"
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"

  Delete "$SMPROGRAMS\${OLD_APP_NAME}\${OLD_APP_NAME}.lnk"
  Delete "$SMPROGRAMS\${OLD_APP_NAME}\Uninstall.lnk"
  RMDir "$SMPROGRAMS\${OLD_APP_NAME}"
  Delete "$DESKTOP\${OLD_APP_NAME}.lnk"

  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "Publisher" "${PUBLISHER_NAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}" "NoRepair" 1

  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\ResearchReportAutomation"
SectionEnd

Section "Uninstall"
  nsExec::ExecToLog 'taskkill /F /T /IM "${APP_EXE}"'
  nsExec::ExecToLog 'taskkill /F /T /IM "${OLD_APP_EXE}"'
  Sleep 1000

  MessageBox MB_YESNO|MB_ICONQUESTION "사용자 데이터도 함께 삭제할까요?$\r$\n$\r$\n실행 기록, API Key 등의 사용자 데이터를 모두 삭제합니다." IDNO skipUserDataDelete
  RMDir /r "${APP_DATA_DIR}"
skipUserDataDelete:

  Delete "$DESKTOP\${APP_NAME}.lnk"
  Delete "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"
  Delete "$SMPROGRAMS\${APP_NAME}\삭제.lnk"
  RMDir "$SMPROGRAMS\${APP_NAME}"

  Delete "$INSTDIR\${APP_EXE}"
  ; 자동 업데이트가 중간에 실패해 남았을 수 있는 백업. 있으면 지우고 없으면 그냥 넘어간다.
  Delete "$INSTDIR\${APP_EXE}.bak"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"

  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${COMPANY_NAME}"
SectionEnd
