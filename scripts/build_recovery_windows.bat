@echo off
chcp 65001 >nul
cd /d "%~dp0.."
if not exist ".venv\Scripts\python.exe" (
 echo Project environment not found. Run scripts\build_windows.bat first or create .venv.
 pause
 exit /b 1
)
".venv\Scripts\python.exe" -m PyInstaller --clean --noconfirm --onefile --windowed --name MaktabOS-Recovery --icon assets\maktabos.ico --paths . scripts\recover_password.py
if errorlevel 1 goto :fail
echo Recovery tool created: dist\MaktabOS-Recovery.exe
pause
exit /b 0
:fail
echo Build failed. See the error above.
pause
exit /b 1
