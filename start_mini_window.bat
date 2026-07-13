@echo off
setlocal
cd /d "%~dp0"
where pythonw.exe >nul 2>nul
if %errorlevel%==0 (
  start "" /b pythonw.exe "%~dp0start_mini_window.py"
) else (
  python "%~dp0start_mini_window.py"
)
endlocal
