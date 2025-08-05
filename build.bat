@echo off
chcp 65001 >nul
title IP管理器 - 快速构建

echo ========================================
echo    IP管理器 - 快速构建工具
echo ========================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python，请先安装Python 3.7+
    pause
    exit /b 1
)

echo 正在检查依赖包...
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo 错误：PyInstaller安装失败
        pause
        exit /b 1
    )
)

echo.
echo 开始构建EXE文件...
echo 这可能需要几分钟时间，请耐心等待...
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
echo 正在复制文件到当前目录...
copy .\dist\IP管理器.exe .\IP管理器.exe >nul 2>&1
if %errorlevel% neq 0 (
    echo 警告：无法复制文件到当前目录
) else (
    echo 已复制到当前目录: IP管理器.exe
)

echo.
echo 构建完成！按任意键退出...
pause