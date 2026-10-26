@echo off
cd /d "%~dp0"
python -m pip install openpyxl
python server.py --port 8880
pause
