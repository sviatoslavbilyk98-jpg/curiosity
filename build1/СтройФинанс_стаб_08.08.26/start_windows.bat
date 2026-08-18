@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 goto run

where python >nul 2>nul
if %errorlevel%==0 goto run

goto missing

:run
py -3 app.py
if %errorlevel% neq 0 python app.py
pause
goto :eof

:missing
echo Python 3 not found.
echo Install Python from python.org and enable "Add Python to PATH".
pause
