; =============================================
; OAS CAHPS Auditor - NSIS Installer Script
; =============================================
; Build with: makensis audit_installer.nsi
; Requires: NSIS 3.x, EnVar plugin
; =============================================

!include "MUI2.nsh"
!include "FileFunc.nsh"
!include "LogicLib.nsh"

!include "nsDialogs.nsh"

; --------------- Configuration ---------------
; VERSION is passed in from package.bat via /DVERSION=x.y.z
; If not provided, default to "0.0.0"
!ifndef VERSION
  !define VERSION "0.0.0"
!endif

!define APPNAME "OAS CAHPS Auditor"
!define PUBLISHER "Tyler Brock"
!define HELPURL "https://github.com/ToonLunk/OAS-CAHPS-Auditor"
!define EXE_NAME "audit.exe"
!define INSTALL_DIR "C:\OAS-CAHPS-Auditor"
!define UNINSTALL_REG_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\OASCAHPSAuditor"

; --------------- General Settings ---------------
Name "${APPNAME} ${VERSION}"
OutFile "..\dist\OAS-CAHPS-Auditor-v${VERSION}-Setup.exe"
InstallDir "${INSTALL_DIR}"
RequestExecutionLevel admin
SetCompressor /SOLID lzma
BrandingText "${APPNAME} v${VERSION}"

; --------------- MUI Settings ---------------
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"
!define MUI_ABORTWARNING
!define MUI_WELCOMEPAGE_TITLE "Welcome to ${APPNAME} Setup"
!define MUI_WELCOMEPAGE_TEXT "This wizard will install ${APPNAME} v${VERSION} on your computer.$\r$\n$\r$\nThe auditor validates OAS CAHPS Excel files and generates HTML reports.$\r$\n$\r$\nClick Next to continue."
!define MUI_FINISHPAGE_TITLE "Installation Complete"
!define MUI_FINISHPAGE_TEXT "${APPNAME} has been installed to:$\r$\n$\r$\n${INSTALL_DIR}$\r$\n$\r$\nYou can now use it from the command line (audit --help) or by right-clicking Excel files and folders in Explorer.$\r$\n$\r$\nNote: You may need to open a new terminal window for the PATH change to take effect."

; --------------- Pages ---------------
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
Page custom SIDsPageCreate SIDsPageLeave
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; --------------- Language ---------------
!insertmacro MUI_LANGUAGE "English"

; --------------- Functions ---------------
Var SIDsCheckbox
Var OpenSIDsPage

Function SIDsPageCreate
  ; Skip this page if SIDs.csv already exists
  IfFileExists "$INSTDIR\SIDs.csv" 0 +2
    Abort

  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  !insertmacro MUI_HEADER_TEXT "SIDs.csv Required" "Download SIDs.csv for SID validation"

  ${NSD_CreateLabel} 0 0 100% 24u "SIDs.csv was not found in the install directory. This file is required for SID validation to work."
  Pop $0

  ${NSD_CreateLabel} 0 30u 100% 24u "Download it from the shared OneDrive folder and save it to:"
  Pop $0

  ${NSD_CreateLabel} 0 58u 100% 12u "$INSTDIR\SIDs.csv"
  Pop $0
  CreateFont $1 "Microsoft Sans Serif" 8 700
  SendMessage $0 ${WM_SETFONT} $1 0

  ${NSD_CreateLabel} 0 78u 100% 24u "See 'About SIDs.csv.txt' in the install folder for details on how to update SIDs.csv when new clients are added."
  Pop $0

  ${NSD_CreateCheckbox} 0 110u 100% 12u "Open the SIDs.csv download page after setup"
  Pop $SIDsCheckbox
  ${NSD_SetState} $SIDsCheckbox ${BST_CHECKED}

  nsDialogs::Show
FunctionEnd

Function SIDsPageLeave
  ${NSD_GetState} $SIDsCheckbox $0
  ${If} $0 == ${BST_CHECKED}
    StrCpy $OpenSIDsPage "1"
  ${Else}
    StrCpy $OpenSIDsPage "0"
  ${EndIf}
FunctionEnd

Function .onGUIEnd
  ${If} $OpenSIDsPage == "1"
    ExecShell "open" "https://jlm353-my.sharepoint.com/:f:/g/personal/dcdata_jlm-solutions_com/IgBhYR7tt6YTRbgNTDEh9M7xAc5HSCC3KSaJt6ImfJV65kg?e=hKp0ZU"
  ${EndIf}
FunctionEnd

; =============================================
; INSTALLER SECTIONS
; =============================================

Section "!Core Files (required)" SecCore
  SectionIn RO ; Read-only, cannot be unchecked

  SetOutPath "$INSTDIR"

  ; Always overwrite the main executable and docs
  File "..\dist\${EXE_NAME}"
  File "..\LICENSE"
  File "..\docs\Installation Instructions.txt"
  File "..\docs\About SIDs.csv.txt"

  ; --- cpt_codes.json with upgrade logic ---
  IfFileExists "$INSTDIR\cpt_codes.json" 0 +3
    MessageBox MB_YESNO|MB_ICONQUESTION "An existing cpt_codes.json was found.$\r$\nDo you want to overwrite it with the new version?" IDYES install_cpt IDNO skip_cpt
    Goto skip_cpt
  install_cpt:
    File "..\cpt_codes.json"
  skip_cpt:

  ; --- Write uninstaller ---
  WriteUninstaller "$INSTDIR\uninstall.exe"

  ; --- Add/Remove Programs entry ---
  WriteRegStr HKLM "${UNINSTALL_REG_KEY}" "DisplayName" "${APPNAME}"
  WriteRegStr HKLM "${UNINSTALL_REG_KEY}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegStr HKLM "${UNINSTALL_REG_KEY}" "QuietUninstallString" '"$INSTDIR\uninstall.exe" /S'
  WriteRegStr HKLM "${UNINSTALL_REG_KEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "${UNINSTALL_REG_KEY}" "Publisher" "${PUBLISHER}"
  WriteRegStr HKLM "${UNINSTALL_REG_KEY}" "URLInfoAbout" "${HELPURL}"
  WriteRegStr HKLM "${UNINSTALL_REG_KEY}" "HelpLink" "${HELPURL}"
  WriteRegStr HKLM "${UNINSTALL_REG_KEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "${UNINSTALL_REG_KEY}" "DisplayIcon" "$INSTDIR\${EXE_NAME},0"
  WriteRegDWORD HKLM "${UNINSTALL_REG_KEY}" "NoModify" 1
  WriteRegDWORD HKLM "${UNINSTALL_REG_KEY}" "NoRepair" 1

  ; --- Estimated size ---
  ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
  IntFmt $0 "0x%08X" $0
  WriteRegDWORD HKLM "${UNINSTALL_REG_KEY}" "EstimatedSize" $0
SectionEnd

Section "Add to System PATH" SecPath
  ; Use EnVar plugin to add install dir to system PATH
  EnVar::SetHKLM
  EnVar::Check "PATH" "$INSTDIR"
  Pop $0
  ${If} $0 != 0
    EnVar::AddValue "PATH" "$INSTDIR"
    Pop $0
  ${EndIf}
SectionEnd

Section "Context Menu Integration" SecContextMenu
  ; --- Folder background: "Audit All OAS Files" (legacy context menu) ---
  WriteRegStr HKCR "Directory\Background\shell\AuditAll" "" "Audit All OAS Files"
  WriteRegStr HKCR "Directory\Background\shell\AuditAll" "Icon" "$INSTDIR\${EXE_NAME},0"
  WriteRegStr HKCR "Directory\Background\shell\AuditAll\command" "" 'cmd.exe /c cd /d "%V" & "$INSTDIR\${EXE_NAME}" --all & pause'

  ; --- Folder background: "Audit All OAS Files" (Windows 11 new context menu) ---
  WriteRegStr HKLM "SOFTWARE\Classes\Directory\Background\shell\AuditAll" "" "Audit All OAS Files"
  WriteRegStr HKLM "SOFTWARE\Classes\Directory\Background\shell\AuditAll" "Icon" "$INSTDIR\${EXE_NAME},0"
  WriteRegStr HKLM "SOFTWARE\Classes\Directory\Background\shell\AuditAll\command" "" 'cmd.exe /c cd /d "%V" & "$INSTDIR\${EXE_NAME}" --all & pause'

  ; --- Individual Excel files: "Audit This OAS File" ---
  ; .xlsx
  WriteRegStr HKCR "SystemFileAssociations\.xlsx\shell\AuditFile" "" "Audit This OAS File"
  WriteRegStr HKCR "SystemFileAssociations\.xlsx\shell\AuditFile" "Icon" "$INSTDIR\${EXE_NAME},0"
  WriteRegStr HKCR "SystemFileAssociations\.xlsx\shell\AuditFile\command" "" 'cmd.exe /c ""$INSTDIR\${EXE_NAME}" "%1" & pause"'

  ; .xls
  WriteRegStr HKCR "SystemFileAssociations\.xls\shell\AuditFile" "" "Audit This OAS File"
  WriteRegStr HKCR "SystemFileAssociations\.xls\shell\AuditFile" "Icon" "$INSTDIR\${EXE_NAME},0"
  WriteRegStr HKCR "SystemFileAssociations\.xls\shell\AuditFile\command" "" 'cmd.exe /c ""$INSTDIR\${EXE_NAME}" "%1" & pause"'

  ; .xlsm
  WriteRegStr HKCR "SystemFileAssociations\.xlsm\shell\AuditFile" "" "Audit This OAS File"
  WriteRegStr HKCR "SystemFileAssociations\.xlsm\shell\AuditFile" "Icon" "$INSTDIR\${EXE_NAME},0"
  WriteRegStr HKCR "SystemFileAssociations\.xlsm\shell\AuditFile\command" "" 'cmd.exe /c ""$INSTDIR\${EXE_NAME}" "%1" & pause"'
SectionEnd

; --------------- Section Descriptions ---------------
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecCore} "Installs the auditor executable and required files."
  !insertmacro MUI_DESCRIPTION_TEXT ${SecPath} "Adds the install directory to the system PATH so you can run 'audit' from any terminal."
  !insertmacro MUI_DESCRIPTION_TEXT ${SecContextMenu} "Adds right-click context menu options to audit Excel files and folders directly from Explorer."
!insertmacro MUI_FUNCTION_DESCRIPTION_END

; =============================================
; UNINSTALLER SECTION
; =============================================

Section "Uninstall"
  ; --- Remove files ---
  Delete "$INSTDIR\${EXE_NAME}"
  Delete "$INSTDIR\LICENSE"
  Delete "$INSTDIR\Installation Instructions.txt"
  Delete "$INSTDIR\About SIDs.csv.txt"
  Delete "$INSTDIR\cpt_codes.json"
  Delete "$INSTDIR\SIDs.csv"
  Delete "$INSTDIR\uninstall.exe"

  ; --- Remove install directory (only if empty) ---
  RMDir "$INSTDIR"

  ; --- Remove from system PATH ---
  EnVar::SetHKLM
  EnVar::DeleteValue "PATH" "$INSTDIR"
  Pop $0

  ; --- Remove context menu entries (legacy) ---
  DeleteRegKey HKCR "Directory\Background\shell\AuditAll"

  ; --- Remove context menu entries (Windows 11) ---
  DeleteRegKey HKLM "SOFTWARE\Classes\Directory\Background\shell\AuditAll"

  ; --- Remove file association context menu entries ---
  DeleteRegKey HKCR "SystemFileAssociations\.xlsx\shell\AuditFile"
  DeleteRegKey HKCR "SystemFileAssociations\.xls\shell\AuditFile"
  DeleteRegKey HKCR "SystemFileAssociations\.xlsm\shell\AuditFile"

  ; --- Remove Add/Remove Programs entry ---
  DeleteRegKey HKLM "${UNINSTALL_REG_KEY}"
SectionEnd
