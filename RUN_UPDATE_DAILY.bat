@echo off
setlocal
cd /d "%~dp0"
title Dashboard Trung Bay - Daily Update

echo ==============================================
echo  UPDATE DAILY - KHONG CAN PYTHON
echo ==============================================
echo.
where powershell.exe >nul 2>nul || (echo LOI: Khong tim thay Windows PowerShell. & pause & exit /b 1)
where git.exe >nul 2>nul || (echo LOI: Khong tim thay Git/GitHub Desktop. & pause & exit /b 1)

echo [1] Build data + tach link NPP...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_daily.ps1"
if errorlevel 1 (echo. & echo BUILD FAILED - xem loi phia tren. & pause & exit /b 1)

echo [2] Commit va push GitHub...
git add docs LINK_NPP.csv source/template.html config.json build_daily.ps1 RUN_UPDATE_DAILY.bat README_VI.md .gitignore
git diff --cached --quiet
if %errorlevel%==0 goto nochange
git commit -m "Daily dashboard update"
if errorlevel 1 (echo COMMIT FAILED & pause & exit /b 1)
git push
if errorlevel 1 (echo PUSH FAILED - mo GitHub Desktop va bam Push origin. & pause & exit /b 1)
goto done

:nochange
echo Khong co thay doi moi de push.
:done
echo.
echo DONE. Mo LINK_NPP.csv de lay link.
start "" "%~dp0LINK_NPP.csv"
pause
