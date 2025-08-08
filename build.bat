@echo off
chcp 65001 >nul
title IP管理器构建工具

echo ========================================
echo    IP管理器 - 构建工具
echo ========================================
echo.

:: 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python
    pause
    exit /b 1
)

:: 检查PyInstaller
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
)

:: 清理旧文件
echo 清理旧文件...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "*.spec" del "*.spec" 2>nul

echo.
echo 开始构建EXE文件...
echo.

:: 构建EXE
pyinstaller --onefile --windowed --name="IP管理器" main.py

if errorlevel 1 (
    echo.
    echo 构建失败！
    pause
    exit /b 1
)

echo.
echo 构建成功！
echo EXE文件位置: dist\IP管理器.exe

:: 复制到当前目录
copy "dist\IP管理器.exe" "IP管理器.exe" >nul 2>&1
if %errorlevel% equ 0 (
    echo 已复制到当前目录: IP管理器.exe
) else (
    echo 复制文件失败
)

echo.
echo 构建完成！
pause 