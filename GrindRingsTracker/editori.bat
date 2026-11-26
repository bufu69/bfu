@echo off
REM Administrare utilizatori (utilizator + parola): utilizator nou, parola noua, administrator,
REM dezactivare, stergere. Util si daca ai uitat parola de administrator.
REM Utilizatorii se pot gestiona si din tracker: butonul "Utilizatori" (doar pentru administratori).
cd /d "%~dp0"
set "PY="
if exist "%~dp0python.txt" set /p PY=<"%~dp0python.txt"
if defined PY if not exist "%PY%" set "PY="
if not defined PY if exist "%~dp0python\python.exe" set "PY=%~dp0python\python.exe"
if not defined PY if exist "%~dp0python.exe" set "PY=%~dp0python.exe"
if not defined PY for /f "delims=" %%E in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%E"
if not defined PY set "PY=python"
"%PY%" server.py --editori
