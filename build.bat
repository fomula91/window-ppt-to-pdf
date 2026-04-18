@echo off
REM PyInstaller 단일 exe 빌드 스크립트
REM 사용: build.bat  (결과: dist\ppt2pdf.exe, dist\doctor.exe)

setlocal

where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo pyinstaller가 설치되어 있지 않습니다. 먼저 아래 명령을 실행하세요:
    echo     pip install pyinstaller
    exit /b 1
)

REM 메인 GUI 앱 (창 없음, 콘솔 숨김)
pyinstaller ^
    --onefile ^
    --windowed ^
    --name ppt2pdf ^
    --clean ^
    app.py
if errorlevel 1 exit /b 1

REM 진단 도구 (콘솔 유지)
pyinstaller ^
    --onefile ^
    --console ^
    --name doctor ^
    --clean ^
    doctor.py
if errorlevel 1 exit /b 1

echo.
echo 빌드 완료:
echo   dist\ppt2pdf.exe  (메인 앱)
echo   dist\doctor.exe   (환경 진단)
endlocal
