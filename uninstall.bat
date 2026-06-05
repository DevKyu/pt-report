@echo off
setlocal

net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo ========================================
echo  PT Report - Uninstall
echo ========================================
echo.

echo Removing scheduled tasks...
for %%t in (PT_Weekly PT_Monthly PT_Weekly_MON_12 PT_Weekly_MON_18 PT_Weekly_TUE_12 PT_Weekly_TUE_18 PT_Weekly_12 PT_Weekly_18 PT_Monthly_10 PT_Monthly_12 PT_Monthly_15 PT_Monthly_18 PT_Monthly_10B PT_Monthly_12B PT_Monthly_15B PT_Monthly_18B) do (
    schtasks /delete /tn "%%t" /f >nul 2>&1
)
echo   Done.

echo.
echo Removing desktop shortcuts...
if exist "%USERPROFILE%\Desktop\PT Report.lnk" (
    del "%USERPROFILE%\Desktop\PT Report.lnk"
    echo   Removed: PT Report.lnk
)
if exist "%USERPROFILE%\Desktop\PT Run Now.bat" (
    del "%USERPROFILE%\Desktop\PT Run Now.bat"
    echo   Removed: PT Run Now.bat
)

echo.
echo ========================================
echo  Uninstall complete.
echo  Please delete the project folder manually.
echo ========================================
echo.
pause
endlocal
