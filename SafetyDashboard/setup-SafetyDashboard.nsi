; Installer pentru Safety Dashboard.
; Compilare cu NSIS (https://nsis.sourceforge.io): click dreapta pe fisier -> "Compile NSIS Script".
; Rezulta setup-SafetyDashboard.exe: copiaza programul in C:\SafetyDashboard si ruleaza
; instalare-server.ps1 (Python, primul editor, firewall, pornire automata).

Unicode true
!define APP "Safety Dashboard"
!define TASK "SafetyDashboard"
!define VERSIUNE "1.0"

Name "${APP}"
OutFile "setup-SafetyDashboard.exe"
InstallDir "C:\SafetyDashboard"
RequestExecutionLevel admin
ShowInstDetails show

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

Section "Program"
  nsExec::Exec 'schtasks /End /TN "${TASK}"'
  SetOutPath "$INSTDIR"
  ; doar fisierele programului - datele existente (safety.db, imagini) raman
  File "server.py"
  File "index.html"
  File "CITESTE.txt"
  File "instalare-server.bat"
  File "instalare-server.ps1"
  File "actualizare-server.bat"
  File "actualizare-server.ps1"
  File "porneste-dashboard.bat"
  File "reporneste-dashboard.bat"
  File "editori.bat"
  File "import-excel.bat"
  File "setup-SafetyDashboard.nsi"
  ; datele initiale (se importa singure la prima pornire), daca fisierul exista langa script
  File /nonfatal "Dashboard safety.xlsm"
  SetOutPath "$INSTDIR\fonts"
  File "fonts\*.*"
  SetOutPath "$INSTDIR\implicit"
  File "implicit\*.*"
  SetOutPath "$INSTDIR"

  WriteUninstaller "$INSTDIR\dezinstalare.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "DisplayName" "${APP}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "DisplayVersion" "${VERSIUNE}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "UninstallString" '"$INSTDIR\dezinstalare.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "InstallLocation" "$INSTDIR"

  CreateDirectory "$SMPROGRAMS\${APP}"
  WriteINIStr "$SMPROGRAMS\${APP}\Dashboard.url" "InternetShortcut" "URL" "http://localhost:8093/"
  CreateShortcut "$SMPROGRAMS\${APP}\Editori.lnk" "$INSTDIR\editori.bat"
  CreateShortcut "$SMPROGRAMS\${APP}\Reporneste dashboard.lnk" "$INSTDIR\reporneste-dashboard.bat"
  CreateShortcut "$SMPROGRAMS\${APP}\Folder program.lnk" "$INSTDIR"

  ExecWait 'powershell -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\instalare-server.ps1"'
SectionEnd

Section "Uninstall"
  nsExec::Exec 'schtasks /End /TN "${TASK}"'
  nsExec::Exec 'schtasks /Delete /F /TN "${TASK}"'
  nsExec::Exec 'schtasks /Delete /F /TN "${TASK}-Mail"'
  nsExec::Exec 'netsh advfirewall firewall delete rule name="Safety Dashboard 8093"'
  Delete "$INSTDIR\server.py"
  Delete "$INSTDIR\index.html"
  Delete "$INSTDIR\CITESTE.txt"
  Delete "$INSTDIR\*.bat"
  Delete "$INSTDIR\*.ps1"
  Delete "$INSTDIR\setup-SafetyDashboard.nsi"
  RMDir /r "$INSTDIR\fonts"
  RMDir /r "$INSTDIR\implicit"
  Delete "$INSTDIR\dezinstalare.exe"
  RMDir /r "$SMPROGRAMS\${APP}"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}"
  MessageBox MB_OK "Programul a fost dezinstalat.$\r$\nDatele au ramas in $INSTDIR (safety.db, imagini, backups)."
SectionEnd
