@echo off
cd /d "%~dp0"
echo Running main.py ...
echo When prompted, paste: C:\Users\Eunice.Lo\Downloads\avcc
echo.
".venv\Scripts\python.exe" main.py
echo.
echo Exit code: %ERRORLEVEL%
pause
