@echo off
REM Cine poate adauga / modifica / sterge actiuni: nume + cod personal.
cd /d "%~dp0"
set "PY="
if exist "%~dp0python.txt" set /p PY=<"%~dp0python.txt"
if defined PY if not exist "%PY%" set "PY="
if not defined PY if exist "%~dp0python\python.exe" set "PY=%~dp0python\python.exe"
if not defined PY if exist "%~dp0python.exe" set "PY=%~dp0python.exe"
if not defined PY for /f "delims=" %%E in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%E"
if not defined PY set "PY=python"
"%PY%" server.py --editori
