@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   Invoice Seq Detector - Build Script
echo ============================================
echo.
echo [1/3] Cleaning old build files...
rmdir /s /q build 2>nul
del /f /q "dist\发票连号检测工具.exe" 2>nul
echo [2/3] Building...
pyinstaller "invoice_seq_detector.spec"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Build failed! Check the error above.
    pause
    exit /b %ERRORLEVEL%
)
echo [3/3] Cleaning build cache...
rmdir /s /q build 2>nul
echo.
echo ============================================
echo   Build completed!
echo   Output: dist\发票连号检测工具.exe
echo ============================================
pause
