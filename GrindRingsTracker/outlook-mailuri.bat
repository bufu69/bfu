@echo off
REM Trimite mailurile cu actiunile de urmarit prin Outlook-ul instalat pe acest calculator.
REM   outlook-mailuri.bat           = trimite
REM   outlook-mailuri.bat /afiseaza = doar deschide mailurile in Outlook, fara sa le trimita
REM   outlook-mailuri.bat /auto     = trimite fara pauza la final (pentru sarcina programata)
cd /d "%~dp0"
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
if not exist mail.ini ( echo Lipseste mail.ini - copiaza mail-exemplu.ini in mail.ini si completeaza-l. & pause & exit /b 1 )
%PY% server.py --raport-mail || ( pause & exit /b 1 )
if /i "%~1"=="/afiseaza" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0trimite-outlook.ps1" -Afiseaza
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0trimite-outlook.ps1"
)
if /i not "%~1"=="/auto" pause
