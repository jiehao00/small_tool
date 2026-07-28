@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   工资条生成工具 - 一键打包
echo ============================================
echo.

:: 1. 关闭正在运行的旧程序
echo [1/3] 关闭旧版程序...
taskkill /f /im "工资条生成工具.exe" >nul 2>&1

:: 2. 删除旧的 EXE
if exist "dist\工资条生成工具.exe" (
    del /f "dist\工资条生成工具.exe" >nul 2>&1
    echo        已删除旧版 EXE
) else (
    echo        无需清理
)

:: 3. 查找 Python
echo.
echo [2/3] 查找 Python 环境...
set PYTHON_CMD=

:: 优先用虚拟环境的 python
if exist "%~dp0venv\Scripts\python.exe" (
    set PYTHON_CMD=%~dp0venv\Scripts\python.exe
)
if exist "%~dp0.venv\Scripts\python.exe" (
    set PYTHON_CMD=%~dp0.venv\Scripts\python.exe
)

:: 否则用系统 python
if "%PYTHON_CMD%"=="" (
    where python >nul 2>&1
    if %errorlevel% equ 0 (
        set PYTHON_CMD=python
    )
)

:: 尝试 py launcher
if "%PYTHON_CMD%"=="" (
    where py >nul 2>&1
    if %errorlevel% equ 0 (
        set PYTHON_CMD=py
    )
)

if "%PYTHON_CMD%"=="" (
    echo [FAIL] 未找到 Python，请先安装 Python 3！
    echo.
    pause
    exit /b 1
)

echo        使用: %PYTHON_CMD%

:: 4. 检查 PyInstaller
echo.
echo [3/3] 开始打包（约 1-2 分钟，请稍候）...
echo.

%PYTHON_CMD% -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] 正在安装 PyInstaller...
    %PYTHON_CMD% -m pip install pyinstaller -q
    if %errorlevel% neq 0 (
        echo [FAIL] PyInstaller 安装失败，请检查网络后重试
        pause
        exit /b 1
    )
)

%PYTHON_CMD% build_exe.py

if %errorlevel% equ 0 (
    if exist "dist\工资条生成工具.exe" (
        echo.
        echo ============================================
        echo   [OK] 打包完成！
        echo   位置: %~dp0dist\工资条生成工具.exe
        echo ============================================
    ) else (
        echo.
        echo [FAIL] EXE 未生成，请查看上方错误信息
    )
) else (
    echo.
    echo [FAIL] 打包失败，请查看上方错误信息
)

echo.
pause
