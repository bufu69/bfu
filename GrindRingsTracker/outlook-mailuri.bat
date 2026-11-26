@echo off
REM Trimite mailurile prin Outlook-ul de pe acest calculator.
REM   outlook-mailuri.bat            trimite
REM   outlook-mailuri.bat /afiseaza  doar deschide mailurile in Outlook (le verifici si apesi Send)
REM   outlook-mailuri.bat /auto      trimite fara pauza (pentru programare zilnica)
cd /d "%~dp0"
set "PY="
if exist "%~dp0python.txt" set /p PY=<"%~dp0python.txt"
if defined PY if not exist "%PY%" set "PY="
if not defined PY if exist "%~dp0python\python.exe" set "PY=%~dp0python\python.exe"
if not defined PY if exist "%~dp0python.exe" set "PY=%~dp0python.exe"
if not defined PY for /f "delims=" %%E in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%E"
if not defined PY set "PY=python"
if not exist mail.ini ( echo Lipseste mail.ini - copiaza mail-exemplu.ini in mail.ini si completeaza-l. & pause & exit /b 1 )
"%PY%" server.py --raport-mail || ( pause & exit /b 1 )
if /i "%~1"=="/afiseaza" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0trimite-outlook.ps1" -Afiseaza
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0trimite-outlook.ps1"
)
if /i not "%~1"=="/auto" pause
