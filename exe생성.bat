@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
echo ╔═══════════════════════════════════════════════════════════════╗
echo ║                                                               ║
echo ║           PDF 이메일 자동발송 EXE 생성 도구                  ║
echo ║                                                               ║
echo ╚═══════════════════════════════════════════════════════════════╝
echo.
echo.

echo [0단계] MAIN_NAME 확인 중...
for /f "tokens=*" %%i in ('python get_main_name.py 2^>nul') do (
    set "EXE_NAME=%%i"
)
if not defined EXE_NAME (
    echo ⚠ MAIN_NAME을 읽을 수 없습니다. 기본값을 사용합니다.
    set "EXE_NAME=PDF 이메일 자동발송"
)
echo ✓ 실행 파일명: !EXE_NAME!
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
    echo ✓ PyInstaller가 이미 설치되어 있습니다
)

echo.
echo Pillow (이미지 처리) 설치 확인 중...
python -c "from PIL import Image" 2>nul
if errorlevel 1 (
    echo ⚠ Pillow가 설치되지 않았습니다. 설치 중...
    pip install Pillow
    echo ✓ Pillow 설치 완료
) else (
    echo ✓ Pillow가 이미 설치되어 있습니다
)
echo.

echo [2단계] 이전 빌드 파일 정리 중...
if exist "build" (
    rmdir /s /q "build"
    echo ✓ build 폴더 삭제 완료
)
if exist "dist" (
    rmdir /s /q "dist"
    echo ✓ dist 폴더 삭제 완료
)
if exist "!EXE_NAME!.spec" (
    del /q "!EXE_NAME!.spec"
    echo ✓ spec 파일 삭제 완료
)
if exist "!EXE_NAME!.exe" (
    del /q "!EXE_NAME!.exe"
    echo ✓ 기존 exe 파일 삭제 완료
)
if exist "!EXE_NAME!.ico" (
    del /q "!EXE_NAME!.ico"
    echo ✓ 기존 ico 파일 삭제 완료
)
echo.

echo [3단계] PNG 아이콘을 ICO로 변환 중...
python -c "from PIL import Image; img = Image.open('PDF 이메일 자동발송_icon.png'); img.save('!EXE_NAME!.ico', format='ICO', sizes=[(256,256)])" 2>nul
if exist "!EXE_NAME!.ico" (
    echo ✓ 아이콘 변환 완료
) else (
    echo ⚠ 아이콘 변환 실패 - Pillow 라이브러리가 없습니다.
    echo   아이콘 없이 진행합니다...
)
echo.

echo [4단계] PyInstaller로 EXE 파일 생성 중...
echo.

if exist "!EXE_NAME!.ico" (
    pyinstaller ^
        --onefile ^
        --windowed ^
        --name "!EXE_NAME!" ^
        --icon="!EXE_NAME!.ico" ^
        --hidden-import tkinter ^
        --hidden-import tkinter.ttk ^
        --hidden-import tkinter.scrolledtext ^
        --hidden-import email ^
        --hidden-import smtplib ^
        --hidden-import pathlib ^
        --hidden-import webbrowser ^
        --hidden-import json ^
        --hidden-import re ^
        --hidden-import copy ^
        pdf_email_sender_gui.py
) else (
    pyinstaller ^
        --onefile ^
        --windowed ^
        --name "!EXE_NAME!" ^
        --hidden-import tkinter ^
        --hidden-import tkinter.ttk ^
        --hidden-import tkinter.scrolledtext ^
        --hidden-import email ^
        --hidden-import smtplib ^
        --hidden-import pathlib ^
        --hidden-import webbrowser ^
        --hidden-import json ^
        --hidden-import re ^
        --hidden-import copy ^
        pdf_email_sender_gui.py
)

echo.
echo.

if exist "dist\!EXE_NAME!.exe" (
    echo [5단계] EXE 파일을 최상위 폴더로 이동 중...
    move /y "dist\!EXE_NAME!.exe" "!EXE_NAME!.exe"
    echo ✓ 이동 완료
    echo.
    
    echo [6단계] 임시 파일 정리 중...
    if exist "build" rmdir /s /q "build"
    if exist "dist" rmdir /s /q "dist"
    if exist "!EXE_NAME!.spec" del /q "!EXE_NAME!.spec"
    if exist "!EXE_NAME!.ico" del /q "!EXE_NAME!.ico"
    echo ✓ 정리 완료
    echo.
    echo.
    
    echo ╔═══════════════════════════════════════════════════════════════╗
    echo ║                                                               ║
    echo ║                  ✅ EXE 파일 생성 완료!                      ║
    echo ║                                                               ║
    echo ║  생성된 파일: !EXE_NAME!.exe                        ║
    echo ║                                                               ║
    echo ║  이제 "!EXE_NAME!.exe" 파일을 더블클릭하면 됩니다!  ║
    echo ║                                                               ║
    echo ║  특징:                                                        ║
    echo ║  - 모든 설정은 프로그램 옆 settings.json에 저장됩니다        ║
    echo ║  - USB에 넣어 이동하면서 사용 가능합니다                     ║
    echo ║  - 폴더는 사용자가 선택할 수 있습니다                        ║
    echo ║  - 사용자 지정 아이콘 적용됨                                 ║
    echo ║                                                               ║
    echo ╚═══════════════════════════════════════════════════════════════╝
) else (
    echo.
    echo ╔═══════════════════════════════════════════════════════════════╗
    echo ║                                                               ║
    echo ║                  ❌ EXE 파일 생성 실패!                      ║
    echo ║                                                               ║
    echo ║  위의 오류 메시지를 확인하세요.                              ║
    echo ║                                                               ║
    echo ╚═══════════════════════════════════════════════════════════════╝
)

echo.
echo.
pause

