@echo off
chcp 65001 >nul
echo Starting hwp with claude installer...
echo (Python will be installed automatically if missing)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
echo.
echo ==========================================
echo   Done. See install_result.txt for details.
echo ==========================================
echo.
echo   Press any key to close...
pause >nul
