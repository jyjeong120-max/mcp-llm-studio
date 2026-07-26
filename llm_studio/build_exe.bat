@echo off
REM LocalLLM Studio exe 빌드 스크립트
REM 사전 조건: pip install -r requirements.txt pyinstaller
chcp 65001 >nul
cd /d "%~dp0"

echo [1/2] PyInstaller 빌드...
pyinstaller --noconfirm LocalLLMStudio.spec
if errorlevel 1 (
  echo 빌드 실패!
  exit /b 1
)

echo [2/2] llama-server 동봉 확인...
if exist llama\llama-server.exe (
  xcopy /e /i /y llama dist\LocalLLMStudio\llama >nul
  echo   llama\ 폴더를 dist에 복사했습니다.
) else (
  echo   [주의] llama\llama-server.exe가 없습니다.
  echo   llama.cpp 릴리스의 win-cuda-x64 빌드를 풀어 llama\ 폴더에 넣으면
  echo   exe와 함께 배포됩니다. (없으면 앱이 목 모드로 시작됩니다)
)

echo.
echo 완료: dist\LocalLLMStudio\LocalLLMStudio.exe
echo 다음 단계: Inno Setup으로 installer.iss를 컴파일하면 Setup.exe가 나옵니다.
