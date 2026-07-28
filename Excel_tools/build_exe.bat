@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================
echo    Excel通用小工具 - 一键打包脚本
echo ============================================
echo.

:: 检查 Python 是否可用
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 并添加到 PATH！
    pause
    exit /b 1
)

:: 检查并安装 PyInstaller
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo [1/4] 正在安装 PyInstaller...
    pip install pyinstaller -q
    echo       安装完成！
) else (
    echo [1/4] PyInstaller 已安装，跳过。
)

:: 安装项目依赖
echo [2/4] 安装项目依赖...
pip install -r requirements.txt -q
echo       依赖安装完成！

:: 清理旧的打包文件
echo [3/4] 清理旧的打包文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "Excel通用小工具.spec" del /q "Excel通用小工具.spec"

:: 开始打包
echo [4/4] 开始打包（请稍候，可能需要几分钟）...
echo.
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "Excel通用小工具" ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.styles ^
    --hidden-import openpyxl.utils ^
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
    echo 输出文件: dist\Excel通用小工具.exe
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
