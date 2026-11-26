@echo off
REM Instaleaza Grind Rings Action Tracker pe acest calculator (ruleaza ca Administrator).
powershell -NoProfile -Command "Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0instalare-server.ps1\"'"
