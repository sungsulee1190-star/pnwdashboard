@echo off
chcp 65001 >nul
setlocal

rem 마켓맵 데이터 배포: build_marketmap.py -> validate_data.py -> 암호화 -> data.enc.js만 commit -> push
rem 피터가 직접 실행한다. 실행 = 업로드 승인(C27). 중간에 실패하면 멈추고 아무것도 올리지 않는다.
cd /d "%~dp0.."

for /f %%b in ('git rev-parse --abbrev-ref HEAD') do set BRANCH=%%b
if not "%BRANCH%"=="main" (
    echo [오류] 지금 브랜치가 main이 아닙니다: %BRANCH%
    pause
    exit /b 1
)

echo [1/4] 마켓맵 데이터 만들기 (data.js에 DATA_MM 넣기)
python pipeline\build_marketmap.py
if errorlevel 1 (
    echo [오류] build_marketmap.py 실패. 업로드하지 않습니다.
    pause
    exit /b 1
)

echo [2/4] data.js 검증 (합계 불일치·지역 미배정 1천톤 초과면 중단)
python validate_data.py
if errorlevel 1 (
    echo [오류] 검증 실패. 원본을 고친 뒤 다시 실행하세요. 업로드하지 않습니다.
    pause
    exit /b 1
)

echo [3/4] 암호화
echo 주의: 입력하는 비밀번호가 화면에 그대로 보입니다. 옆에 사람이 없을 때 입력하세요.
echo       비밀번호에 ^& ^| ^< ^> ^^ %% 기호가 있으면 이 창에서 깨질 수 있습니다. 그때는 node tools\encrypt-data.mjs를 직접 쓰세요.
set /p "PNW_PASSWORD=대시보드 비밀번호: "
set /p "PNW_CONFIRM=한 번 더 입력: "
if not "%PNW_PASSWORD%"=="%PNW_CONFIRM%" (
    echo [오류] 두 번 입력한 비밀번호가 다릅니다. 업로드하지 않습니다.
    pause
    exit /b 1
)
set PNW_CONFIRM=
cls
node tools\encrypt-data.mjs
if errorlevel 1 (
    set PNW_PASSWORD=
    echo [오류] 암호화 실패. 업로드하지 않습니다.
    pause
    exit /b 1
)
set PNW_PASSWORD=

echo [4/4] data.enc.js만 올리기 (평문 data.js는 올리지 않음)
git add data.enc.js
git commit -m "data: 마켓맵 데이터(DATA_MM) 갱신 (data.enc.js)"
if errorlevel 1 (
    echo [오류] commit 실패(바뀐 내용이 없을 수 있음).
    pause
    exit /b 1
)
git push origin main
if errorlevel 1 (
    echo [오류] push 실패. 강제 push는 하지 않습니다. git pull --rebase origin main 후 다시 실행하세요.
    pause
    exit /b 1
)

echo.
echo 배포 완료. 1~3분 뒤 사이트에서 Ctrl+F5 후 2페이지 '마켓맵'을 확인하세요.
pause
