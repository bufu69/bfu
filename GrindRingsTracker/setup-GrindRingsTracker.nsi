; Installer pentru Grind Rings Action Tracker.
; Se compileaza cu NSIS (https://nsis.sourceforge.io): click dreapta pe acest fisier -> "Compile NSIS Script".
; Rezulta setup-GrindRingsTracker.exe, care copiaza fisierele in C:\GrindRingsTracker
; si ruleaza instalare-server.ps1 (Python, editori, firewall, pornire automata).

Unicode true
!define APPNAME "Grind Rings Action Tracker"
!define TASK "GrindRingsTracker"
!define VERSION "1.0"

Name "${APPNAME}"
OutFile "setup-GrindRingsTracker.exe"
InstallDir "C:\GrindRingsTracker"
RequestExecutionLevel admin
ShowInstDetails show

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

Section "Instalare"
  SetOutPath "$INSTDIR"
  ; programul (datele existente - grind_rings.db, attachments, mail.ini - nu se suprascriu)
  File "server.py"
  File "index.html"
  File "grafice.html"
  File "pagina-veche-past_due_tracker.html"
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
  SetOutPath "$INSTDIR\fonts"
  File "fonts\*.*"
  SetOutPath "$INSTDIR"

  ; opreste serverul daca ruleaza deja (actualizare)
  nsExec::Exec 'schtasks /End /TN "${TASK}"'

  WriteUninstaller "$INSTDIR\dezinstalare.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "DisplayName" "${APPNAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "UninstallString" '"$INSTDIR\dezinstalare.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}" "InstallLocation" "$INSTDIR"

  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\Dashboard.lnk" "http://localhost:8092/"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\Editori.lnk" "$INSTDIR\editori.bat"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\Reporneste tracker.lnk" "$INSTDIR\reporneste-tracker.bat"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\Folder.lnk" "$INSTDIR"

  ; configurarea (Python, primul editor, firewall, pornire automata)
  ExecWait 'powershell -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\instalare-server.ps1"'
SectionEnd

Section "Uninstall"
  nsExec::Exec 'schtasks /End /TN "${TASK}"'
  nsExec::Exec 'schtasks /Delete /F /TN "${TASK}"'
  nsExec::Exec 'schtasks /Delete /F /TN "${TASK}-Mail"'
  nsExec::Exec 'netsh advfirewall firewall delete rule name="Grind Rings Action Tracker 8092"'
  Delete "$INSTDIR\server.py"
  Delete "$INSTDIR\*.html"
  Delete "$INSTDIR\*.bat"
  Delete "$INSTDIR\*.ps1"
  Delete "$INSTDIR\CITESTE.txt"
  Delete "$INSTDIR\mail-exemplu.ini"
  Delete "$INSTDIR\setup-GrindRingsTracker.nsi"
  RMDir /r "$INSTDIR\fonts"
  RMDir /r "$INSTDIR\mail_out"
  Delete "$INSTDIR\dezinstalare.exe"
  RMDir /r "$SMPROGRAMS\${APPNAME}"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${TASK}"
  ; baza de date, atasamentele, backup-urile si mail.ini raman in $INSTDIR
  MessageBox MB_OK "Programul a fost dezinstalat. Datele (grind_rings.db, attachments, backups, mail.ini) au ramas in $INSTDIR."
SectionEnd
