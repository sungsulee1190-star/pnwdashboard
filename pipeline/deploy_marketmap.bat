@echo off
rem Marketmap deploy: build_marketmap.py -> validate_data.py -> encrypt -> commit data.enc.js only -> push
rem Peter runs this himself. Running it = upload approval (C27). Any failure stops before upload.
rem This file stays ASCII-only: cmd misreads multibyte lines after "chcp 65001". Python/Node print Korean.
chcp 65001 >nul
setlocal
cd /d "%~dp0.."

for /f %%b in ('git rev-parse --abbrev-ref HEAD') do set BRANCH=%%b
if not "%BRANCH%"=="main" (
    echo [ERROR] Current branch is not main: %BRANCH%
    pause
    exit /b 1
)

echo [1/4] Build marketmap data (DATA_MM into data.js)
python pipeline\build_marketmap.py
if errorlevel 1 goto fail

echo [2/4] Validate data.js (stops if mismatch or unassigned region over 1,000 t)
python validate_data.py
if errorlevel 1 goto fail

echo [3/4] Encrypt
echo WARNING: the password is visible on screen while you type. Make sure nobody is watching.
echo          If the password has special characters such as ^& ^| ^< ^> ^^ %%, use node tools\encrypt-data.mjs instead.
set "PNW_PASSWORD="
set "PNW_CONFIRM="
set /p "PNW_PASSWORD=Dashboard password: "
set /p "PNW_CONFIRM=Type it again: "
if not defined PNW_PASSWORD goto mismatch
if not "%PNW_PASSWORD%"=="%PNW_CONFIRM%" goto mismatch
set "PNW_CONFIRM="
cls
node tools\encrypt-data.mjs
if errorlevel 1 goto fail
set "PNW_PASSWORD="

echo [4/4] Commit data.enc.js only (plain data.js is never uploaded)
git add data.enc.js
git commit -m "data: update marketmap data (DATA_MM) in data.enc.js"
if errorlevel 1 goto fail
git push origin main
if errorlevel 1 (
    echo [ERROR] Push failed. No force push. Run "git pull --rebase origin main" and try again.
    pause
    exit /b 1
)

echo.
echo Done. Wait 1-3 minutes, press Ctrl+F5 on the site, open page 2 and choose the marketmap mode.
pause
exit /b 0

:mismatch
set "PNW_PASSWORD="
set "PNW_CONFIRM="
echo [ERROR] The two passwords are empty or different. Nothing was uploaded.
pause
exit /b 1

:fail
set "PNW_PASSWORD="
echo [ERROR] A step failed (see the message above). Nothing was uploaded.
pause
exit /b 1
