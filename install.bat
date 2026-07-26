@echo off
chcp 65001 >nul
set PYTHONUTF8=1
echo ==========================================
echo   hwp with claude - Installer
echo ==========================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    py "%~dp0install.py"
    goto :end
)

where python >nul 2>nul
if %errorlevel%==0 (
    python "%~dp0install.py"
    goto :end
)

echo [!] Python not found.
echo     Install Python first: https://www.python.org/downloads/
echo     During install, CHECK "Add Python to PATH".
echo     Then run this file again.
echo.
pause

:end
