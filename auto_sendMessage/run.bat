@echo off
chcp 65001 >nul
title 企业微信自动发送工具

python -c "import pyautogui" >nul 2>&1
if %errorlevel% neq 0 (
    echo 首次运行，正在安装依赖...
    pip install pyautogui pygetwindow pyperclip -i https://pypi.tuna.tsinghua.edu.cn/simple
    echo.
)

python main.py
pause
