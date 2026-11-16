@echo off
cd /d "%~dp0"
echo ========================================
echo  autocad-layer-combine
echo  Script: %CD%\main.py
echo  Python: %CD%\.venv\Scripts\python.exe
echo  You MUST see: Build: com-retry-4
echo ========================================
echo When prompted, paste: C:\Users\Eunice.Lo\Downloads\avcc
echo.
".venv\Scripts\python.exe" -u "%~dp0main.py"
echo.
echo Exit code: %ERRORLEVEL%
pause
