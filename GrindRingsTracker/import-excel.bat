@echo off
REM Importa actiuni dintr-un Excel (.xlsx / .xlsm). Trage fisierul peste acest .bat
REM sau scrie calea cand te intreaba. Coloanele se recunosc dupa nume (vezi CITESTE.txt).
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
set "FOAIE="
set /p "FOAIE=Foaia (Enter = o gasesc automat): "
if "%FOAIE%"=="" ( "%PY%" server.py --import-excel "%FISIER%" ) else ( "%PY%" server.py --import-excel "%FISIER%" --foaie "%FOAIE%" )
pause
