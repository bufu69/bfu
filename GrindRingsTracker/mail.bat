@echo off
REM Trimite mailurile (actiuni intarziate / aproape de termen / urgente) prin SMTP, cu setarile din mail.ini.
REM Pentru Outlook foloseste outlook-mailuri.bat.   mail.bat /auto = fara pauza la final.
cd /d "%~dp0"
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
if not exist mail.ini ( echo Lipseste mail.ini - copiaza mail-exemplu.ini in mail.ini si completeaza-l. & pause & exit /b 1 )
%PY% server.py --trimite-mail
if /i not "%~1"=="/auto" pause
