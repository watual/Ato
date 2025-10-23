@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ================================================================
echo.
echo           PDF 이메일 자동발송 프로그램 런처
echo.
echo ================================================================
echo.
echo [필요 파일]
echo - pdf_email_sender_gui.py (같은 폴더에 필요)
echo - favicon\favicon.ico (아이콘 교체용, 없으면 기본 아이콘 사용)
echo.
echo.

echo [0단계] MAIN_NAME 확인 중...
for /f "tokens=*" %%i in ('python pdf_email_sender_gui.py --get-main-name 2^>nul') do (
    set "EXE_NAME=%%i"
)
if not defined EXE_NAME (
    echo ❌ Python 실행 오류 또는 MAIN_NAME을 읽을 수 없습니다.
    echo.
    echo Python이 설치되어 있는지 확인하거나 기본값을 사용합니다.
    set "EXE_NAME=PDF 이메일 자동발송"
)
echo ✓ 실행 파일명: !EXE_NAME!
echo.

echo ================================================================
echo.
echo           실행 모드를 선택해주세요
echo.
echo   1. EXE 파일 생성 (PyInstaller 필요)
echo   2. GUI 프로그램 직접 실행 (Python 필요)
echo   3. 종료
echo.
echo ================================================================
echo.

choice /c 123 /d 1 /t 30 /n /m "선택 (1-3, 30초 후 자동으로 1번 선택): "
set SELECTED=%ERRORLEVEL%

if "%SELECTED%"=="1" goto build_exe
if "%SELECTED%"=="2" goto run_gui
if "%SELECTED%"=="3" goto end_wait
goto end


:run_gui
echo.
echo ================================================================
echo.
echo           GUI 프로그램 직접 실행
echo.
echo ================================================================
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
    goto end
) else (
    echo ✓ Python이 설치되어 있습니다
)

echo.
echo [2단계] 필요한 라이브러리 설치 확인 중...
echo.

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
    goto end
) else (
    echo ✓ smtplib 사용 가능
)

echo.
echo [3단계] 프로그램 실행 중...
echo.
echo ================================================================
echo.
echo              PDF 이메일 자동발송 프로그램 시작
echo.
echo ================================================================
echo.

python pdf_email_sender_gui.py

echo.
echo ================================================================
echo.
echo              프로그램이 종료되었습니다
echo.
echo ================================================================
echo.
goto end_wait


:build_exe
echo.
echo ================================================================
echo.
echo           EXE 파일 생성
echo.
echo ================================================================
echo.

echo [1단계] 필요한 라이브러리 설치 확인 중...
echo.
echo PyInstaller 설치 확인 중...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo ⚠ PyInstaller가 설치되지 않았습니다. 설치 중...
    pip install pyinstaller
    echo ✓ PyInstaller 설치 완료
) else (
    echo ✓ PyInstaller 사용 가능
)
echo.

echo [2단계] 기존 파일 정리 중...
if exist "!EXE_NAME!.exe" (
    del /q "!EXE_NAME!.exe"
    echo ✓ 기존 exe 파일 삭제 완료
)
echo.

echo [3단계] PyInstaller로 EXE 파일 생성 중...
echo.

if exist "favicon\favicon.ico" (
    pyinstaller --onefile --windowed --name "!EXE_NAME!" --icon="favicon\favicon.ico" --hidden-import tkinter --hidden-import tkinter.ttk --hidden-import tkinter.scrolledtext --hidden-import tkinter.messagebox --hidden-import tkinter.filedialog --hidden-import email.mime.text --hidden-import email.mime.multipart --hidden-import email.mime.application --hidden-import smtplib --hidden-import socket --hidden-import json --hidden-import re --hidden-import datetime --hidden-import pathlib --hidden-import logging --hidden-import threading --hidden-import webbrowser --hidden-import copy --hidden-import time --hidden-import os --hidden-import sys --add-data "favicon;favicon" pdf_email_sender_gui.py
) else (
    echo ⚠ favicon.ico 파일을 찾을 수 없습니다. 아이콘 없이 진행합니다...
    pyinstaller --onefile --windowed --name "!EXE_NAME!" --hidden-import tkinter --hidden-import tkinter.ttk --hidden-import tkinter.scrolledtext --hidden-import tkinter.messagebox --hidden-import tkinter.filedialog --hidden-import email.mime.text --hidden-import email.mime.multipart --hidden-import email.mime.application --hidden-import smtplib --hidden-import socket --hidden-import json --hidden-import re --hidden-import datetime --hidden-import pathlib --hidden-import logging --hidden-import threading --hidden-import webbrowser --hidden-import copy --hidden-import time --hidden-import os --hidden-import sys --add-data "favicon;favicon" pdf_email_sender_gui.py
)

if errorlevel 1 (
    echo.
    echo ❌ PyInstaller 실행 중 오류 발생!
    echo 위의 오류 메시지를 확인하세요.
    pause
    goto end
)

echo.
echo [5단계] EXE 파일 이동 중...
if exist "dist\!EXE_NAME!.exe" (
    move "dist\!EXE_NAME!.exe" "!EXE_NAME!.exe" >nul
    echo ✓ EXE 파일 이동 완료: !EXE_NAME!.exe
) else (
    echo ❌ EXE 파일을 찾을 수 없습니다.
    echo dist 폴더에 !EXE_NAME!.exe 파일이 생성되지 않았습니다.
    pause
    goto end
)
echo.

echo [6단계] 파일 정리 중...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "!EXE_NAME!.spec" del /q "!EXE_NAME!.spec"
echo ✓ 정리 완료
echo.

echo ================================================================
echo.
echo              EXE 파일 생성이 완료되었습니다!
echo.
echo  - 실행 파일: !EXE_NAME!.exe
echo  - USB에 넣어 이동하면서 사용 가능합니다
echo  - 폴더는 사용자가 선택할 수 있습니다
echo  - 사용자 지정 아이콘 적용됨
echo.
echo ================================================================
echo.
goto end_wait


:end_wait
echo.
echo 3초 후 자동으로 종료됩니다...
echo (Enter / P / p = 일시정지, 그 외 키 = 즉시 종료, 입력 없음 = 자동 종료)
echo.

powershell -NoProfile -Command "$end=(Get-Date).AddSeconds(3); while((Get-Date) -lt $end){ if([Console]::KeyAvailable){ $k=[Console]::ReadKey($true); if($k.Key -eq 'Enter' -or $k.KeyChar -in 'p','P'){ exit 20 } else { exit 30 } } Start-Sleep -Milliseconds 50 }; exit 0"
set "rc=%errorlevel%"

if "%rc%"=="20" (
    echo.
    echo [일시정지] Enter 또는 P/p 입력. 로그를 확인하세요.
    echo 아무 키나 누르면 종료...
    pause >nul
    goto end
) else if "%rc%"=="30" (
    echo.
    echo [즉시종료] Enter/P/p 이외의 키 입력.
    goto end
) else (
    echo.
    echo [자동종료] 3초 경과로 종료합니다.
    goto end
)


:end
exit /b 0