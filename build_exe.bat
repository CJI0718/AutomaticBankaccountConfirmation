@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  금융기관조회서 변환기 - 실행파일(.exe) 빌드
echo ============================================
echo.
echo [1/2] 필요 패키지 설치 중...
pip install -e . pyinstaller
if errorlevel 1 goto :err
echo.
echo [2/2] 실행 파일 빌드 중... (수 분 소요)
pyinstaller --noconfirm --onefile --windowed ^
  --name 금융기관조회서변환기 ^
  --add-data "configs;configs" ^
  --collect-all pymupdf ^
  afc\gui.py
if errorlevel 1 goto :err
echo.
echo ============================================
echo  완료!  dist\금융기관조회서변환기.exe
echo  이 exe 하나만 회계사 PC에 복사해서 더블클릭하면 됩니다.
echo ============================================
pause
exit /b 0

:err
echo.
echo [실패] 위 메시지를 확인하세요. (인터넷 연결/파이썬 설치 여부 확인)
pause
exit /b 1
