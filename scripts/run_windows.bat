@echo off
chcp 65001 >nul
cd /d "%~dp0.."
where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (set "PY=py -3") else (set "PY=python")
if not exist ".venv\Scripts\python.exe" (
 echo [1/2] Creating Python virtual environment...
 %PY% -m venv .venv || goto :fail
)
echo [2/2] Installing dependencies and starting MaktabOS...
".venv\Scripts\python.exe" -m pip install -r requirements.txt || goto :fail
".venv\Scripts\python.exe" -m schoolregister
if errorlevel 1 goto :fail
exit /b 0
:fail
echo Startup failed. Check that Python 3.10+ is installed and internet is available for dependency installation.
pause
exit /b 1
