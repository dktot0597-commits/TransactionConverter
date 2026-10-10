@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    set "PY=py -3"
) else (
    set "PY=python"
)
%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installing openpyxl ...
    %PY% -m pip install --user openpyxl
    if errorlevel 1 goto failed
)
%PY% -c "import tkinterdnd2" >nul 2>nul
if errorlevel 1 (
    echo Installing tkinterdnd2 for drag and drop ...
    %PY% -m pip install --user tkinterdnd2
    if errorlevel 1 goto failed
)
%PY% "TransactionConverter_GUI_v4_4.py"
if errorlevel 1 pause
exit /b 0
:failed
echo [ERROR] Dependency installation failed. Check internet connection and Python settings.
pause
exit /b 1
