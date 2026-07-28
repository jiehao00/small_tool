@echo off
chcp 65001 >nul
title 打包exe - 企业微信自动发送工具

echo ================================
echo   企业微信自动发送工具 - 打包exe
echo ================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

echo [1/5] 关闭旧进程...
taskkill /f /im "企业微信自动发送工具.exe" 2>nul
echo   (如有旧进程已关闭)

echo.
echo [2/5] 安装依赖...
pip install pyautogui pygetwindow pyperclip pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
if %errorlevel% neq 0 (
    echo [错误] 安装失败
    pause
    exit /b 1
)

echo.
echo [3/5] 打包（约需 1-3 分钟）...
pyinstaller --onefile ^
    --noconsole ^
    --name "企业微信自动发送工具" ^
    --hidden-import pyautogui ^
    --hidden-import pygetwindow ^
    --hidden-import pyperclip ^
    --hidden-import cv2 ^
    --clean ^
    --noconfirm ^
    main.py

if %errorlevel% neq 0 (
    echo.
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo [4/5] 复制配置文件...
if exist "config.json" (
    copy /Y "config.json" "dist\config.json" >nul
    echo   config.json 已复制到 dist\
)

echo.
echo [5/5] 复制截图资源...
if exist "images" (
    xcopy /E /I /Y "images" "dist\images" >nul
    echo   images\ 已复制到 dist\
)

echo.
echo ================================
echo   打包完成！
echo   exe: dist\企业微信自动发送工具.exe
echo ================================
echo.
echo 使用说明:
echo   1. 首次运行前，先点"坐标校准"获取屏幕坐标
echo   2. 填入文件夹路径和群名称
echo   3. 点击"开始发送"
echo ================================
pause
