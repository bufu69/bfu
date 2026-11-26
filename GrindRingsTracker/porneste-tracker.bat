@echo off
REM Porneste Grind Rings Action Tracker (port 8092).
REM Instalat (instalare-server.bat): porneste serverul in fundal.
REM Neinstalat: il porneste in aceasta fereastra (se opreste cand inchizi fereastra).
cd /d "%~dp0"
schtasks /Query /TN "GrindRingsTracker" >nul 2>&1
if %errorlevel%==0 (
  schtasks /Run /TN "GrindRingsTracker" >nul
  echo Serverul ruleaza in fundal.
  timeout /t 3 >nul
  start "" "http://localhost:8092/"
  exit /b 0
)
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
echo Pornesc serverul... (inchide fereastra ca sa-l opresti)
start "" cmd /c "timeout /t 3 >nul & start http://localhost:8092/"
%PY% server.py
pause
