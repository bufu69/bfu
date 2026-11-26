@echo off
cd /d "%~dp0"
python -m pip install --quiet openpyxl
python grind_rings_server.py
pause
