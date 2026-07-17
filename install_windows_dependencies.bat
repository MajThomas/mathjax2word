@echo off
cd /d "%~dp0"
winget install python
python -m pip install -r requirements.txt
winget install Inkscape.inkscape
pause
