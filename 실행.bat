@echo off
chcp 65001 > nul
echo ╔═══════════════════════════════════════════════════════════════╗
echo ║                                                               ║
echo ║           PDF 이메일 자동발송 프로그램 실행                  ║
echo ║                                                               ║
echo ╚═══════════════════════════════════════════════════════════════╝
echo.
echo.

echo [1단계] Python 설치 확인 중...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python이 설치되지 않았습니다.
    echo.
    echo Python을 설치해주세요:
    echo https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
) else (
    echo ✓ Python이 설치되어 있습니다
)

echo.
echo [2단계] 필요한 라이브러리 설치 확인 중...
echo.
echo 필요한 라이브러리들을 확인하고 설치합니다...

python -c "import tkinter" 2>nul
if errorlevel 1 (
    echo ⚠ tkinter가 없습니다. Python 재설치가 필요할 수 있습니다.
) else (
    echo ✓ tkinter 사용 가능
)

python -c "import smtplib" 2>nul
if errorlevel 1 (
    echo ❌ smtplib가 없습니다. Python 설치에 문제가 있습니다.
    pause
    exit /b 1
) else (
    echo ✓ smtplib 사용 가능
)

echo.
echo [3단계] 설정 파일 확인 중...
if not exist "settings.json" (
    echo ❌ settings.json 파일이 없습니다.
    echo.
    echo 설정 파일을 먼저 생성해주세요.
    echo GUI 버전을 실행하여 설정을 완료하세요.
    echo.
    pause
    exit /b 1
) else (
    echo ✓ settings.json 파일 확인됨
)

echo.
echo [4단계] 전송할PDF 폴더 확인 중...
if not exist "전송할PDF" (
    echo ⚠ 전송할PDF 폴더가 없습니다. 생성합니다...
    mkdir "전송할PDF"
    echo ✓ 전송할PDF 폴더 생성 완료
) else (
    echo ✓ 전송할PDF 폴더 확인됨
)

if not exist "전송완료" (
    echo ⚠ 전송완료 폴더가 없습니다. 생성합니다...
    mkdir "전송완료"
    echo ✓ 전송완료 폴더 생성 완료
) else (
    echo ✓ 전송완료 폴더 확인됨
)

echo.
echo [5단계] 프로그램 실행 중...
echo.
echo ╔═══════════════════════════════════════════════════════════════╗
echo ║                                                               ║
echo ║              PDF 이메일 자동발송 프로그램 시작              ║
echo ║                                                               ║
echo ╚═══════════════════════════════════════════════════════════════╝
echo.

python pdf_email_sender.py

echo.
echo ╔═══════════════════════════════════════════════════════════════╗
echo ║                                                               ║
echo ║              프로그램이 종료되었습니다                       ║
echo ║                                                               ║
echo ╚═══════════════════════════════════════════════════════════════╝
echo.
pause
