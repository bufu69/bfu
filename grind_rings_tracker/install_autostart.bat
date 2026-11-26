@echo off
REM Porneste automat Grind Rings Action Tracker la pornirea calculatorului
REM (Task Scheduler, fara sa fie nevoie de utilizator logat). Ruleaza ca Administrator.
cd /d "%~dp0"
for /f "delims=" %%P in ('where pythonw 2^>nul') do if not defined PYW set "PYW=%%P"
if not defined PYW for /f "delims=" %%P in ('where python') do if not defined PYW set "PYW=%%P"
if not defined PYW ( echo Python nu a fost gasit in PATH. & pause & exit /b 1 )
schtasks /Create /F /TN "GrindRingsActionTracker" /SC ONSTART /RU SYSTEM /RL HIGHEST ^
  /TR "\"%PYW%\" \"%~dp0grind_rings_server.py\""
REM Deschide portul 8092 in firewall
netsh advfirewall firewall add rule name="Grind Rings Action Tracker 8092" dir=in action=allow protocol=TCP localport=8092 >nul
schtasks /Run /TN "GrindRingsActionTracker"
echo Gata. Serverul porneste automat si ruleaza acum pe http://%COMPUTERNAME%:8092/
pause
