@echo off
REM Cine poate adauga / modifica / sterge actiuni: nume + cod personal.
cd /d "%~dp0"
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
%PY% server.py --editori
