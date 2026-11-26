@echo off
REM Importa datele din "Dashboard safety.xlsm" (Good Catch, Echipamente, Actiuni corective, Incidente).
REM Trage fisierul Excel peste acest .bat sau scrie calea cand te intreaba.
REM Atentie: implicit INLOCUIESTE datele existente (se face backup inainte).
cd /d "%~dp0"
set "PY="
if exist "%~dp0python.txt" set /p PY=<"%~dp0python.txt"
if defined PY if not exist "%PY%" set "PY="
if not defined PY if exist "%~dp0python\python.exe" set "PY=%~dp0python\python.exe"
if not defined PY if exist "%~dp0python.exe" set "PY=%~dp0python.exe"
if not defined PY for /f "delims=" %%E in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%E"
if not defined PY set "PY=python"
set "FISIER=%~1"
if "%FISIER%"=="" set /p "FISIER=Fisierul Excel (il poti trage aici): "
set "FISIER=%FISIER:"=%"
if not exist "%FISIER%" ( echo Nu exista: %FISIER% & pause & exit /b 1 )
set "MOD="
set /p "MOD=Inlocuiesti toate datele cu cele din Excel? (D = inlocuieste, N = adauga) [D]: "
if /i "%MOD%"=="N" ( "%PY%" server.py --import-excel "%FISIER%" --adauga ) else ( "%PY%" server.py --import-excel "%FISIER%" )
pause
