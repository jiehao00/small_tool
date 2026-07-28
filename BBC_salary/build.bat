@echo off
chcp 65001 >nul
title 工资汇总明细合并工具 - 一键打包

echo ============================================
echo   工资汇总明细合并工具 - 一键打包
echo ============================================
echo.

rem 清理旧文件
echo [1/3] 清理旧构建文件...
if exist "build" rmdir /s /q "build"
if exist "dist" del /f /q "dist\工资汇总明细合并工具.exe" 2>nul
echo       清理完成！

echo.
echo [2/3] 语法检查...
python -c "import py_compile; py_compile.compile('combine_salary.py', doraise=True); print('       语法检查通过！')"
if %errorlevel% neq 0 (
    echo       语法检查失败，请检查代码！
    pause
    exit /b 1
)

echo.
echo [3/3] 开始打包（可能需要1-2分钟，请耐心等待）...
pyinstaller --onefile --noconsole --name "工资汇总明细合并工具" --clean --noconfirm combine_salary.py

if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo   打包成功！
    echo   输出文件: dist\工资汇总明细合并工具.exe
    echo ============================================
) else (
    echo.
    echo ============================================
    echo   打包失败，请检查上方错误信息！
    echo ============================================
)

echo.
pause
