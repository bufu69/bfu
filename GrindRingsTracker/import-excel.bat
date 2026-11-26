@echo off
REM Importa actiuni dintr-un fisier Excel (.xlsx / .xlsm) in baza de date.
REM Poti trage fisierul Excel peste acest .bat sau il alegi cand te intreaba.
cd /d "%~dp0"
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
set "FISIER=%~1"
if "%FISIER%"=="" set /p "FISIER=Calea catre fisierul Excel (poti trage fisierul aici): "
set "FISIER=%FISIER:"=%"
if not exist "%FISIER%" ( echo Fisierul nu exista: %FISIER% & pause & exit /b 1 )
set "FOAIE="
set /p "FOAIE=Numele foii (Enter = se detecteaza automat): "
if "%FOAIE%"=="" ( %PY% server.py --import-excel "%FISIER%" ) else ( %PY% server.py --import-excel "%FISIER%" --foaie "%FOAIE%" )
echo.
echo Actiunile importate apar in dashboard la urmatoarea reimprospatare.
pause
