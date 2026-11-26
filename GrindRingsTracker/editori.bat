@echo off
REM Administrare editori: cine poate adauga / modifica / sterge actiuni (cod personal).
cd /d "%~dp0"
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
%PY% server.py --editori
