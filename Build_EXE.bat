@echo off
setlocal EnableExtensions
cd /d "%~dp0"
chcp 65001 >nul
if not exist "%~dp0TransactionConverter_GUI_v4_4.py" (
    echo [ERROR] Missing TransactionConverter_GUI_v4_4.py next to this BAT.
    pause
    exit /b 1
)
where py >nul 2>nul
if not errorlevel 1 (
    py -3 "%~dp0build_windows.py"
) else (
    python "%~dp0build_windows.py"
)
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. Read the error above.
    pause
    exit /b 1
)
echo.
echo [OK] EXE created in dist folder.
pause
