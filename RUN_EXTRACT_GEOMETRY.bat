@echo off
cd /d "%~dp0"
echo ========================================
echo  3D polyline geometry CSV export
echo  Script: %CD%\extract_3d_polyline_geometry.py
echo  Python: %CD%\.venv\Scripts\python.exe
echo ========================================
echo.
".venv\Scripts\python.exe" -u "%~dp0extract_3d_polyline_geometry.py"
echo.
echo Exit code: %ERRORLEVEL%
pause
