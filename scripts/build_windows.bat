@echo off
chcp 65001 >nul
cd /d "%~dp0.."
where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (set "PY=py -3") else (set "PY=python")
if not exist ".venv\Scripts\python.exe" (
 %PY% -m venv .venv || goto :fail
)
".venv\Scripts\python.exe" -m pip install -r requirements-build.txt || goto :fail
".venv\Scripts\python.exe" -m PyInstaller --clean --noconfirm --windowed --onedir --name MaktabOS --icon assets\maktabos.ico --add-data "assets;assets" --collect-submodules openpyxl --collect-submodules docx --collect-submodules cryptography launcher.py
if errorlevel 1 goto :fail
set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
 echo Inno Setup 6 is required to create the installer.
 echo Download it from https://jrsoftware.org/isdl.php, then run this script again.
 goto :fail
)
"%ISCC%" installer\MaktabOS.iss || goto :fail
echo Installer created: dist\MaktabOS-Setup-3.0.0.exe
pause
exit /b 0
:fail
echo Build failed. See the error above.
pause
exit /b 1
