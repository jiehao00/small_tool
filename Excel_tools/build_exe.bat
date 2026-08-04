@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================
echo    财务Excel工具箱 - 一键打包脚本
echo ============================================
echo.

:: 使用 Python 3.11
set PYTHON=C:\Users\41147\.workbuddy\binaries\python\versions\3.11.9\python.exe
set PIP=%PYTHON% -m pip

:: 检查 Python 是否可用
%PYTHON% --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python 3.11，请检查路径！
    pause
    exit /b 1
)

:: 检查并安装 PyInstaller
%PIP% show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo [1/4] 正在安装 PyInstaller...
    %PIP% install pyinstaller -q
    echo       安装完成！
) else (
    echo [1/4] PyInstaller 已安装，跳过。
)

:: 安装项目依赖
echo [2/4] 安装项目依赖...
%PIP% install -r requirements.txt -q
echo       依赖安装完成！

:: 清理旧的打包文件
echo [3/4] 清理旧的打包文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "财务Excel工具箱.spec" del /q "财务Excel工具箱.spec"

:: 开始打包
echo [4/4] 开始打包（请稍候，可能需要几分钟）...
echo.
%PYTHON% -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "财务Excel工具箱" ^
    --icon "ChatGPT-Image-2026年7月28日-14_44_50.ico" ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.styles ^
    --hidden-import openpyxl.utils ^
    --hidden-import matplotlib ^
    --hidden-import matplotlib.backends.backend_agg ^
    --hidden-import fpdf2 ^
    --hidden-import fpdf ^
    --add-data "pages;pages" ^
    --clean ^
    --noconfirm ^
    excel_tool.py

if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo    打包成功！
    echo ============================================
    echo.
    echo 输出文件: dist\财务Excel工具箱.exe
    echo.
    :: 打开输出目录
    start "" "dist"
) else (
    echo.
    echo ============================================
    echo    打包失败，请检查上方错误信息！
    echo ============================================
)

echo.
pause
