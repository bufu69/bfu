@echo off
REM Importa actiuni dintr-un Excel (.xlsx / .xlsm). Trage fisierul peste acest .bat
REM sau scrie calea cand te intreaba. Coloanele se recunosc dupa nume (vezi CITESTE.txt).
cd /d "%~dp0"
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
set "FISIER=%~1"
if "%FISIER%"=="" set /p "FISIER=Fisierul Excel (il poti trage aici): "
set "FISIER=%FISIER:"=%"
if not exist "%FISIER%" ( echo Nu exista: %FISIER% & pause & exit /b 1 )
set "FOAIE="
set /p "FOAIE=Foaia (Enter = o gasesc automat): "
if "%FOAIE%"=="" ( %PY% server.py --import-excel "%FISIER%" ) else ( %PY% server.py --import-excel "%FISIER%" --foaie "%FOAIE%" )
pause
