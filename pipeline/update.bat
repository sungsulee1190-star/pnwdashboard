@echo off
chcp 65001 >nul
setlocal

rem 저장소 루트(pnwdashboard-main)에서 실행한다. 어디서 더블클릭해도 같다.
cd /d "%~dp0.."
set LOG_FILE=pipeline\update_log.txt

echo ==============================================
echo [PNW Dashboard 자동 업데이트 프로세스 시작]
echo ==============================================
echo [1/5] 데이터 파싱 및 data.js 생성 (run_updates.py)
python pipeline\run_updates.py --no-git
if errorlevel 1 (
    echo [오류] run_updates.py 실행 실패
    echo %date% %time% - ERROR: run_updates.py 실행 실패 >> "%LOG_FILE%"
    pause
    exit /b 1
)

echo [2/5] Build marketmap data (DATA_MM)
python pipeline\build_marketmap.py
if errorlevel 1 (
    echo [ERROR] Marketmap build failed. Upload cancelled.
    echo %date% %time% - ERROR: Marketmap build failed >> "%LOG_FILE%"
    pause
    exit /b 1
)

echo [3/5] data.js 문법 및 데이터 무결성 검증 (validate_data.py)
python validate_data.py
if errorlevel 1 (
    echo [오류] data.js 검증 실패. 안전을 위해 Git Push를 취소합니다.
    echo %date% %time% - ERROR: data.js 검증 실패 >> "%LOG_FILE%"
    pause
    exit /b 1
)

echo [4/5] data.js 암호화 및 GitHub 배포 준비 (data.js 평문은 올리지 않음)
if not defined PNW_PASSWORD (
    echo [오류] PNW_PASSWORD 환경변수가 없습니다. 대시보드 로그인 비밀번호로 한 번 설정하세요: setx PNW_PASSWORD "비밀번호"
    echo %date% %time% - ERROR: PNW_PASSWORD 미설정 >> "%LOG_FILE%"
    pause
    exit /b 1
)
node tools\encrypt-data.mjs
if errorlevel 1 (
    echo [오류] data.enc.js 암호화 실패
    echo %date% %time% - ERROR: 암호화 실패 >> "%LOG_FILE%"
    pause
    exit /b 1
)
git add data.enc.js index.html dashboard_logic.js
git commit -m "Auto update data.enc.js %date%"

echo [5/5] GitHub로 데이터 전송 중...
git push origin main
if errorlevel 1 (
    echo [경고] Push 실패. GitHub의 새 커밋을 먼저 받아온 뒤 다시 시도합니다...
    git pull --rebase origin main
    git push origin main
    if errorlevel 1 (
        echo [오류] GitHub Push 실패. 강제 Push는 GitHub의 다른 작업을 지우므로 하지 않습니다. 수동 확인 필요.
        echo %date% %time% - ERROR: GitHub Push 실패 >> "%LOG_FILE%"
        pause
        exit /b 1
    )
)

echo.
echo ✅ 모든 업데이트 및 배포가 완료되었습니다.
echo %date% %time% - SUCCESS: 자동 업데이트 성공 >> "%LOG_FILE%"
echo 변경사항이 웹페이지에 반영되기까지 약 1~3분 정도 소요될 수 있습니다.
pause
