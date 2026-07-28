@echo off
chcp 65001 >nul
echo ================================
echo   进项数据采集票种校验工具 - 打包
echo ================================
echo.
pyinstaller "进项数据采集票种校验工具.spec" --noconfirm --clean
echo.
if %errorlevel% equ 0 (
    echo 打包成功！exe 位置: dist\进项数据采集票种校验工具.exe
) else (
    echo 打包失败，请检查错误信息。
)
pause
