@echo off
REM Actualizeaza un Safety Dashboard deja instalat cu fisierele din ACEST folder (versiunea noua).
REM Datele (baza de date, imagini, backups) nu se ating.
powershell -NoProfile -Command "Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0actualizare-server.ps1\"'"
