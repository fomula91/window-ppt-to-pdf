@echo off
REM PyInstaller 단일 exe 빌드 스크립트
REM 사용: build.bat  (빌드 결과는 dist\ppt2pdf.exe)

setlocal

where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo pyinstaller가 설치되어 있지 않습니다. 먼저 아래 명령을 실행하세요:
    echo     pip install pyinstaller
    exit /b 1
)

pyinstaller ^
    --onefile ^
    --windowed ^
    --name ppt2pdf ^
    --clean ^
    app.py

echo.
echo 빌드 완료: dist\ppt2pdf.exe
endlocal
