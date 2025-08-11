@echo off
chcp 65001 >nul
title IP Manager Build Tool

echo ========================================
echo    IP Manager - Build Tool
echo ========================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found
    pause
    exit /b 1
)

:: Clean old files
echo Cleaning old files...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "*.spec" del "*.spec" 2>nul

echo.
echo Starting EXE build...
echo.

:: Build command
set CMD=pyinstaller --noconfirm --clean --onefile --windowed --name="IP Manager"

if exist "ip_manager.ico" (
    set CMD=%CMD% --icon="ip_manager.ico"
)
if exist "ip_manager_256x256.png" (
    set CMD=%CMD% --add-data="ip_manager_256x256.png;."
)
if exist "LibreHardwareMonitor" (
    set CMD=%CMD% --add-data="LibreHardwareMonitor;LibreHardwareMonitor"
)

set CMD=%CMD% --hidden-import=clr --hidden-import=System --hidden-import=System.Threading --hidden-import=System.Collections --hidden-import=win32timezone --hidden-import=pystray --hidden-import=PIL --hidden-import=PIL.Image --hidden-import=PIL.ImageDraw

echo Executing command: %CMD%
%CMD% main.py

if errorlevel 1 (
  echo.
  echo Build failed!
  pause
  exit /b 1
)

if exist "dist\IP Manager.exe" (
  copy /Y "dist\IP Manager.exe" "IP Manager.exe" >nul 2>&1
  echo Copied to current directory: IP Manager.exe
)

echo.
echo Build successful!
echo EXE location: dist\IP Manager.exe
echo.
echo Build completed!
pause 