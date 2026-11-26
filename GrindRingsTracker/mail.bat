@echo off
REM Trimite acum mailurile (actiuni intarziate / aproape de termen / urgente) prin SMTP.
REM Setarile (server, destinatari, trimitere automata) se fac in tracker: Setari -> Email.
REM Pentru Outlook foloseste outlook-mailuri.bat.   mail.bat /auto = fara pauza la final.
cd /d "%~dp0"
set "PY="
if exist "%~dp0python.txt" set /p PY=<"%~dp0python.txt"
if defined PY if not exist "%PY%" set "PY="
if not defined PY if exist "%~dp0python\python.exe" set "PY=%~dp0python\python.exe"
if not defined PY if exist "%~dp0python.exe" set "PY=%~dp0python.exe"
if not defined PY for /f "delims=" %%E in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%E"
if not defined PY set "PY=python"
"%PY%" server.py --trimite-mail
if /i not "%~1"=="/auto" pause
