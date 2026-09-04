@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ================================================
echo   신규입사자 면담 예약 - 서버와 공개 주소 켜기
echo ================================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [오류] Node.js 를 찾지 못했습니다.
  echo        https://nodejs.org 에서 LTS 버전을 설치한 뒤 다시 실행해주세요.
  pause
  exit /b 1
)

echo [1/2] 예약 서버를 시작합니다...
start "면담예약 서버" cmd /c "node server.js & pause"

REM 서버가 포트를 잡을 때까지 잠깐 기다립니다.
timeout /t 3 /nobreak >nul

where cloudflared >nul 2>nul
if errorlevel 1 (
  echo.
  echo [알림] cloudflared 가 없어 공개 주소는 만들지 않습니다.
  echo        이 PC에서만 http://localhost:3000 으로 접속하실 수 있습니다.
  echo.
  echo        공개 주소가 필요하면 아래를 실행한 뒤 이 파일을 다시 실행해주세요.
  echo        winget install Cloudflare.cloudflared
  echo.
  pause
  exit /b 0
)

echo [2/2] 공개 주소를 만듭니다. 아래에 나오는 trycloudflare.com 주소를 신규입사자에게 안내하세요.
echo.
echo        이 창을 닫으면 공개 주소가 사라집니다. 면담 기간 동안 켜두세요.
echo.

cloudflared tunnel --url http://localhost:3000
