@echo off
REM Opreste si elimina pornirea automata. Ruleaza ca Administrator.
schtasks /End /TN "GrindRingsActionTracker" >nul 2>&1
schtasks /Delete /F /TN "GrindRingsActionTracker"
netsh advfirewall firewall delete rule name="Grind Rings Action Tracker 8092" >nul
pause
