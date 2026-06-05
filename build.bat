@echo off
setlocal

set "DIR=%~dp0"
set "DIR=%DIR:~0,-1%"
set "DIST=%DIR%\dist\pt_report_dist"

echo.
echo ========================================
echo  PT Report - EXE Build
echo ========================================
echo.

echo [1/4] Installing dependencies...
python -m pip install --quiet --upgrade pandas openpyxl xlrd apscheduler workalendar
if %errorLevel% neq 0 (
    echo ERROR: pip install failed. Check internet connection.
    pause & exit /b 1
)
python -m PyInstaller --version >nul 2>&1
if %errorLevel% neq 0 (
    python -m pip install pyinstaller --quiet
    if %errorLevel% neq 0 (
        echo ERROR: PyInstaller install failed.
        pause & exit /b 1
    )
)
echo Done.
echo.

echo [2/4] Building exe... (1-3 min)
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "ptreport" ^
    --add-data "%DIR%\scripts;scripts" ^
    --add-data "%DIR%\run.py;." ^
    --add-data "%DIR%\run_silent.vbs;." ^
    --hidden-import pandas ^
    --hidden-import pandas.io.excel._openpyxl ^
    --hidden-import pandas.io.excel._xlrd ^
    --collect-all openpyxl ^
    --collect-all workalendar ^
    --hidden-import convertdate ^
    --hidden-import apscheduler ^
    --hidden-import process ^
    --hidden-import excel_report ^
    --hidden-import html_dashboard ^
    --hidden-import validator ^
    --hidden-import logging.handlers ^
    --hidden-import xlrd ^
    --distpath "%DIST%" ^
    --workpath "%DIR%\build_tmp" ^
    --specpath "%DIR%\build_tmp" ^
    --noconfirm ^
    "%DIR%\app.py"

if %errorLevel% neq 0 (
    echo.
    echo ERROR: Build failed.
    pause & exit /b 1
)
echo Done.
echo.

echo [3/4] Setting up dist folder...
if not exist "%DIST%\data\weekly"  mkdir "%DIST%\data\weekly"
if not exist "%DIST%\data\monthly" mkdir "%DIST%\data\monthly"
if not exist "%DIST%\output"       mkdir "%DIST%\output"
copy "%DIR%\run_silent.vbs" "%DIST%\" >nul
copy "%DIR%\README.txt"     "%DIST%\" >nul
copy "%DIR%\uninstall.bat"  "%DIST%\" >nul
echo Done.
echo.

rd /s /q "%DIR%\build_tmp" >nul 2>&1

echo [4/4] Done.
echo.
echo ========================================
echo  Build complete!
echo  Output: dist\pt_report_dist\
echo ========================================
echo.
echo  ptreport.exe       - Double-click to run GUI
echo  run_silent.vbs     - Called by Task Scheduler (--auto mode)
echo  uninstall.bat      - Remove scheduled tasks
echo  data\weekly\       - Put weekly xlsx here
echo  data\monthly\      - Put monthly xlsx here
echo  output\            - Results saved here
echo.
echo  Zip dist\pt_report_dist\ folder to distribute.
echo.

explorer "%DIST%"
pause
endlocal
