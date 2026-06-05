@echo off
setlocal EnableDelayedExpansion

set "DIR=%~dp0"
set "DIR=%DIR:~0,-1%"
set "VBS=%DIR%\run_silent.vbs"

echo.
echo ========================================
echo  PT Report - Setup
echo ========================================
echo  Path: %DIR%
echo.

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting admin rights...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo [1/4] Checking Python...
python --version >nul 2>&1
if %errorLevel% equ 0 (
    for /f "tokens=*" %%i in ('python --version') do echo   Found: %%i
    goto :python_ok
)

echo   Python not found. Trying winget...
winget --version >nul 2>&1
if %errorLevel% equ 0 (
    echo   Installing Python 3.12...
    winget install Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
    python --version >nul 2>&1
    if %errorLevel% equ 0 goto :python_ok
)

echo   Auto-install failed.
echo   Please install Python from: https://www.python.org/downloads/
echo   IMPORTANT: Check "Add Python to PATH" during install.
start https://www.python.org/downloads/
pause & exit /b 1

:python_ok
echo.

echo [2/4] Installing packages...
python -m pip install --upgrade pip --quiet
python -m pip install pandas openpyxl apscheduler workalendar xlrd --quiet --upgrade
if %errorLevel% neq 0 (
    echo   ERROR: Package install failed. Check internet connection.
    pause & exit /b 1
)
echo   Done.
echo.

echo [3/4] Creating folders...
if not exist "%DIR%\data\weekly"  mkdir "%DIR%\data\weekly"
if not exist "%DIR%\data\monthly" mkdir "%DIR%\data\monthly"
if not exist "%DIR%\output"       mkdir "%DIR%\output"
echo   Done.
echo.

echo [4/4] Registering task scheduler...

:: Write PowerShell script to temp file (locale-independent DaysOfWeek enum)
set "PS1=%TEMP%\pt_reg.ps1"
(
  echo $vbs = '%VBS%'
  echo $old = @('PT_Weekly','PT_Monthly','PT_Weekly_MON_12','PT_Weekly_MON_18','PT_Weekly_TUE_12','PT_Weekly_TUE_18','PT_Weekly_12','PT_Weekly_18','PT_Monthly_10','PT_Monthly_12','PT_Monthly_15','PT_Monthly_18','PT_Monthly_10B','PT_Monthly_12B','PT_Monthly_15B','PT_Monthly_18B'^)
  echo foreach ^($n in $old^) { Unregister-ScheduledTask -TaskName $n -Confirm:$false -ErrorAction SilentlyContinue }
  echo $act = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ^('"' + $vbs + '"'^)
  echo $twk = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At '12:00'
  echo $tmo = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At '12:00'
  echo $cfg = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ^(New-TimeSpan -Minutes 30^) -MultipleInstances IgnoreNew
  echo $pri = New-ScheduledTaskPrincipal -UserId $Env:USERNAME -LogonType Interactive -RunLevel Limited
  echo Register-ScheduledTask -TaskName 'PT_Weekly'  -Action $act -Trigger $twk -Settings $cfg -Principal $pri -Force ^| Out-Null
  echo Register-ScheduledTask -TaskName 'PT_Monthly' -Action $act -Trigger $tmo -Settings $cfg -Principal $pri -Force ^| Out-Null
) > "%PS1%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
del "%PS1%" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -Command "if (Get-ScheduledTask -TaskName 'PT_Weekly' -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }" >nul 2>&1
if %errorLevel% equ 0 (
    echo   Weekly  : Monday at 12:00
    echo   Monthly : 1st at 12:00
    echo   Status  : OK
) else (
    echo   WARNING: Registration failed. Try running as Administrator.
)
echo.

:: Create shortcut on Desktop for manual run
set "SC_RUN=%USERPROFILE%\Desktop\PT Run Now.bat"
(
echo @echo off
echo cd /d "%DIR%"
echo python run.py --force
echo pause
) > "%SC_RUN%"

echo   Shortcut created: "PT Run Now.bat" on Desktop
echo.

echo ========================================
echo  Setup complete!
echo ========================================
echo.
echo  Usage:
echo    Weekly  - Put xlsx in: data\weekly\
echo    Monthly - Put xlsx in: data\monthly\
echo    Auto    - Runs in background automatically
echo.
echo  Manual run (when needed):
echo    Double-click "PT Run Now.bat" on Desktop
echo.
echo  Check log: run.log
echo.
pause
endlocal
