; Installer pentru Grind Rings Action Tracker.
; Compilare cu NSIS (https://nsis.sourceforge.io): click dreapta pe fisier -> "Compile NSIS Script".
; Rezulta setup-GrindRingsTracker.exe: copiaza programul in C:\GrindRingsTracker si ruleaza
; instalare-server.ps1 (Python, primul editor, firewall, pornire automata).

Unicode true
!define APP "Grind Rings Action Tracker"
!define TASK "GrindRingsTracker"
!define VERSIUNE "1.0"

Name "${APP}"
OutFile "setup-GrindRingsTracker.exe"
InstallDir "C:\GrindRingsTracker"
RequestExecutionLevel admin
ShowInstDetails show

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

Section "Program"
  nsExec::Exec 'schtasks /End /TN "${TASK}"'
  SetOutPath "$INSTDIR"
  ; doar fisierele programului - datele existente (grind_rings.db, atasamente, mail.ini) raman
  File "server.py"
  File "index.html"
  File "grafice.html"
  File "CITESTE.txt"
  File "mail-exemplu.ini"
  File "instalare-server.bat"
  File "instalare-server.ps1"
  File "actualizare-server.bat"
  File "actualizare-server.ps1"
  File "porneste-tracker.bat"
  File "reporneste-tracker.bat"
  File "editori.bat"
  File "import-excel.bat"
  File "mail.bat"
  File "outlook-mailuri.bat"
  File "trimite-outlook.ps1"
  File "setup-GrindRingsTracker.nsi"

  WriteUninstaller "$INSTDIR\dezinstalare.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "DisplayName" "${APP}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "DisplayVersion" "${VERSIUNE}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "UninstallString" '"$INSTDIR\dezinstalare.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "InstallLocation" "$INSTDIR"

  CreateDirectory "$SMPROGRAMS\${APP}"
  WriteINIStr "$SMPROGRAMS\${APP}\Tracker.url" "InternetShortcut" "URL" "http://localhost:8092/"
  WriteINIStr "$SMPROGRAMS\${APP}\Grafice.url" "InternetShortcut" "URL" "http://localhost:8092/grafice"
  CreateShortcut "$SMPROGRAMS\${APP}\Editori.lnk" "$INSTDIR\editori.bat"
  CreateShortcut "$SMPROGRAMS\${APP}\Reporneste tracker.lnk" "$INSTDIR\reporneste-tracker.bat"
  CreateShortcut "$SMPROGRAMS\${APP}\Folder program.lnk" "$INSTDIR"

  ExecWait 'powershell -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\instalare-server.ps1"'
SectionEnd

Section "Uninstall"
  nsExec::Exec 'schtasks /End /TN "${TASK}"'
  nsExec::Exec 'schtasks /Delete /F /TN "${TASK}"'
  nsExec::Exec 'schtasks /Delete /F /TN "${TASK}-Mail"'
  nsExec::Exec 'netsh advfirewall firewall delete rule name="Grind Rings Action Tracker 8092"'
  Delete "$INSTDIR\server.py"
  Delete "$INSTDIR\index.html"
  Delete "$INSTDIR\grafice.html"
  Delete "$INSTDIR\CITESTE.txt"
  Delete "$INSTDIR\mail-exemplu.ini"
  Delete "$INSTDIR\*.bat"
  Delete "$INSTDIR\*.ps1"
  Delete "$INSTDIR\setup-GrindRingsTracker.nsi"
  RMDir /r "$INSTDIR\mail_out"
  Delete "$INSTDIR\dezinstalare.exe"
  RMDir /r "$SMPROGRAMS\${APP}"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}"
  MessageBox MB_OK "Programul a fost dezinstalat.$\r$\nDatele au ramas in $INSTDIR (grind_rings.db, atasamente, backups, mail.ini)."
SectionEnd
