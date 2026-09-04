@echo off
chcp 65001 >nul
cd /d "%~dp0"

where node >nul 2>&1
if errorlevel 1 (
  echo.
  echo  Node.js 가 없습니다. https://nodejs.org 에서 LTS 를 설치한 뒤
  echo  이 파일을 다시 더블클릭하세요.
  echo.
  pause
  exit /b 1
)

if not exist "node_modules\" (
  echo 처음 실행입니다. 필요한 구성을 설치합니다...
  call npm install
  if errorlevel 1 (
    echo 설치에 실패했습니다.
    pause
    exit /b 1
  )
)

echo.
echo  브라우저가 곧 열립니다. 이 검은 창은 닫지 마세요.
echo.
node server.js
echo.
pause
