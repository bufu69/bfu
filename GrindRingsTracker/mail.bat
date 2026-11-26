@echo off
REM Trimite mailurile (actiuni intarziate / aproape de termen / urgente) prin SMTP, cu setarile din mail.ini.
REM Pentru Outlook foloseste outlook-mailuri.bat.   mail.bat /auto = fara pauza la final.
cd /d "%~dp0"
set "PY="
if exist "%~dp0python.txt" set /p PY=<"%~dp0python.txt"
if defined PY if not exist "%PY%" set "PY="
if not defined PY if exist "%~dp0python\python.exe" set "PY=%~dp0python\python.exe"
if not defined PY if exist "%~dp0python.exe" set "PY=%~dp0python.exe"
if not defined PY for /f "delims=" %%E in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%E"
if not defined PY set "PY=python"
if not exist mail.ini ( echo Lipseste mail.ini - copiaza mail-exemplu.ini in mail.ini si completeaza-l. & pause & exit /b 1 )
"%PY%" server.py --trimite-mail
if /i not "%~1"=="/auto" pause
