@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Creating .venv ...
  python -m venv .venv
)
echo Installing requirements into .venv only ...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo Install failed.
  exit /b 1
)
echo.
echo Done. Activate with: .venv\Scripts\activate.bat
echo Then run: python main.py
pause
