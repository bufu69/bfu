@echo off
REM Porneste Safety Dashboard (port 8093).
REM Instalat (instalare-server.bat): porneste serverul in fundal.
REM Neinstalat: il porneste in aceasta fereastra (se opreste cand inchizi fereastra).
cd /d "%~dp0"
schtasks /Query /TN "SafetyDashboard" >nul 2>&1
if %errorlevel%==0 (
  schtasks /Run /TN "SafetyDashboard" >nul
  echo Serverul ruleaza in fundal.
  timeout /t 3 >nul
  start "" "http://localhost:8093/"
  exit /b 0
)
set "PY="
if exist "%~dp0python.txt" set /p PY=<"%~dp0python.txt"
if defined PY if not exist "%PY%" set "PY="
if not defined PY if exist "%~dp0python\python.exe" set "PY=%~dp0python\python.exe"
if not defined PY if exist "%~dp0python.exe" set "PY=%~dp0python.exe"
if not defined PY for /f "delims=" %%E in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%E"
if not defined PY set "PY=python"
echo Pornesc serverul... (inchide fereastra ca sa-l opresti)
start "" cmd /c "timeout /t 3 >nul & start http://localhost:8093/"
"%PY%" server.py
pause
