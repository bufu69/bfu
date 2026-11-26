@echo off
REM Trimite acum mailurile cu actiunile intarziate / aproape de termen / urgente,
REM prin serverul de mail (SMTP) din mail.ini. Pentru Outlook foloseste outlook-mailuri.bat.
cd /d "%~dp0"
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
if not exist mail.ini ( echo Lipseste mail.ini - copiaza mail-exemplu.ini in mail.ini si completeaza-l. & pause & exit /b 1 )
%PY% server.py --trimite-mail
if /i not "%~1"=="/auto" pause
