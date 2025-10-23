@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

set "PATH=%PATH%;C:\Windows\System32\WindowsPowerShell\v1.0"
mode con codepage select=65001 >nul 2>&1

set "CERT_SUBJECT=CN=MySelfSignedCodeSign"
set "PFX_PATH=%~dp0selfsign_codesign.pfx"
set "PFX_PASS=1234"
set "TIMESTAMP_URL=http://timestamp.sectigo.com"
set "TRUST_LOCAL_ROOT=0"

echo [1/5] signtool.exe 찾는 중...

set "SIGNSRC="
for %%P in (
  "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe"
  "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x64\signtool.exe"
  "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x64\signtool.exe"
  "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"
) do (
  if exist %%~fP set "SIGNSRC=%%~fP" & goto :signtool_found
)

:: PATH에 있으면 where로 탐색
for /f "usebackq delims=" %%A in (`where signtool 2^>NUL`) do (
  if exist "%%~fA" set "SIGNSRC=%%~fA" & goto :signtool_found
)

:signtool_found
if not defined SIGNSRC (
  echo 오류: signtool.exe를 찾지 못했습니다. Windows SDK를 설치하시거나 PATH를 확인하세요.
  exit /b 1
)
echo    -> "%SIGNSRC%"

echo [2/5] 자체 서명 인증서 존재 여부 확인/생성 중...

:: PowerShell을 사용하여 인증서 생성(없을 때만) + PFX 내보내기
"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass ^
  " $sub='%CERT_SUBJECT%';" ^
  " $pfx='%PFX_PATH%';" ^
  " $pwd=ConvertTo-SecureString '%PFX_PASS%' -AsPlainText -Force;" ^
  " $cert=(Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $sub } | Select-Object -First 1);" ^
  " if(-not $cert) {" ^
  "   $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $sub -KeyExportPolicy Exportable -KeyUsage DigitalSignature -CertStoreLocation 'Cert:\CurrentUser\My';" ^
  " }" ^
  " if(Test-Path $pfx) { Remove-Item -Force $pfx }" ^
  " Export-PfxCertificate -Cert $cert -FilePath $pfx -Password $pwd | Out-Null"

if errorlevel 1 (
  echo 오류: 인증서 생성/내보내기에 실패했습니다.
  :: 임시 파일 정리
  if exist "%PFX_PATH%" del /q "%PFX_PATH%" 2>nul
  if exist "%~dp0대상" del /q "%~dp0대상" 2>nul
  if exist "%~dp0서명" del /q "%~dp0서명" 2>nul
  exit /b 1
)

if "%TRUST_LOCAL_ROOT%"=="1" (
  echo [선택] 로컬 신뢰 루트에 자체서명 인증서 등록 중...
  "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass ^
    " $sub='%CERT_SUBJECT%';" ^
    " $cert=(Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $sub } | Select-Object -First 1);" ^
    " if($cert) { Import-Certificate -FilePath ($cert.PSPath) -CertStoreLocation 'Cert:\CurrentUser\Root' ^| Out-Null }"
)

echo [3/5] 서명 대상 파일 수집 중...

set "TARGETS="

:: 인자로 파일이 들어오면 그 파일만
if not "%~1"=="" (
  if exist "%~1" (
    set "TARGETS=%~1"
  ) else (
    echo 오류: 지정한 파일을 찾을 수 없습니다: %~1
    exit /b 1
  )
) else (
  :: 현재 폴더의 *.exe
  for %%F in ("%~dp0*.exe") do (
    if exist "%%~fF" set "TARGETS=!TARGETS!|%%~fF"
  )
  :: .\dist 폴더의 *.exe
  if exist "%~dp0dist" (
    for %%F in ("%~dp0dist\*.exe") do (
      if exist "%%~fF" set "TARGETS=!TARGETS!|%%~fF"
    )
  )
)

if "%TARGETS%"=="" (
  echo 오류: 서명할 대상 .exe 파일이 없습니다. 인자로 exe 경로를 넘기거나, 현재/.\dist 폴더에 exe를 두세요.
  exit /b 1
)

:: 파이프 구분자 제거
set "TARGETS=%TARGETS:~1%"

echo    -> 대상:
for %%F in (%TARGETS:|= %) do echo       %%~fF

echo [4/5] 서명 진행...
for %%F in (%TARGETS:|= %) do (
  echo    -> 서명: "%%~fF"
  "%SIGNSRC%" sign ^
    /f "%PFX_PATH%" ^
    /p "%PFX_PASS%" ^
    /fd sha256 ^
    /td sha256 ^
    /tr "%TIMESTAMP_URL%" ^
    "%%~fF"
  if errorlevel 1 (
    echo      실패: %%~fF
    :: 임시 파일 정리
    if exist "%PFX_PATH%" del /q "%PFX_PATH%" 2>nul
    if exist "%~dp0대상" del /q "%~dp0대상" 2>nul
    if exist "%~dp0서명" del /q "%~dp0서명" 2>nul
    goto :end
  )
)

echo [5/5] 서명 확인...
set "SIGN_SUCCESS=0"
for %%F in (%TARGETS:|= %) do (
  echo    -> 서명 확인: %%~fF
  for /f "tokens=*" %%S in ('powershell -NoProfile -Command "(Get-AuthenticodeSignature '%%~fF').Status"') do set "STATUS=%%S"
  
  if "!STATUS!"=="Valid" (
    echo       ✅ 공인 서명
    set "SIGN_SUCCESS=1"
  ) else if "!STATUS!"=="UnknownError" (
    echo       ✅ 자체 서명
    set "SIGN_SUCCESS=1"
  ) else if "!STATUS!"=="NotSigned" (
    echo       ❌ 서명 없음
  ) else if "!STATUS!"=="HashMismatch" (
    echo       ❌ 파일 변조됨
  ) else (
    echo       ⚠️ 알 수 없는 상태: !STATUS!
  )
)

if "%SIGN_SUCCESS%"=="1" (
  echo.
  echo ================================================================
  echo [SUCCESS] 자체 서명 완료
  echo ================================================================
  :: 임시 파일 정리
  if exist "%PFX_PATH%" del /q "%PFX_PATH%" 2>nul
  if exist "%~dp0대상" del /q "%~dp0대상" 2>nul
  if exist "%~dp0서명" del /q "%~dp0서명" 2>nul
) else (
  echo.
  echo ================================================================
  echo [FAILED] 서명 실패
  echo ================================================================
  :: 임시 파일 정리
  if exist "%PFX_PATH%" del /q "%PFX_PATH%" 2>nul
  if exist "%~dp0대상" del /q "%~dp0대상" 2>nul
  if exist "%~dp0서명" del /q "%~dp0서명" 2>nul
)
goto :end

:end
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
    goto :exit
) else if "%rc%"=="30" (
    echo.
    echo [즉시종료] Enter/P/p 이외의 키 입력.
    goto :exit
) else (
    echo.
    echo [자동종료] 3초 경과로 종료합니다.
    goto :exit
)


:exit
exit /b 0
