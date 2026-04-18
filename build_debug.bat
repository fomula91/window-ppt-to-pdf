@echo off
REM 디버그(콘솔) 빌드: --windowed 대신 --console 로 stderr/stdout 을 보여준다.
REM 실행은 반드시 cmd.exe 에서 `dist\ppt2pdf_debug.exe` 로 해야 콘솔 창이 유지된다.
REM doctor.exe 도 함께 빌드한다.

setlocal

where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo pyinstaller가 설치되어 있지 않습니다. 먼저 아래 명령을 실행하세요:
    echo     pip install pyinstaller
    exit /b 1
)

pyinstaller ^
    --onefile ^
    --console ^
    --name ppt2pdf_debug ^
    --clean ^
    --debug=noarchive ^
    app.py
if errorlevel 1 exit /b 1

pyinstaller ^
    --onefile ^
    --console ^
    --name doctor ^
    --clean ^
    doctor.py
if errorlevel 1 exit /b 1

echo.
echo 디버그 빌드 완료:
echo   dist\ppt2pdf_debug.exe  (콘솔 모드 메인 앱)
echo   dist\doctor.exe         (환경 진단)
echo 실행 방법: cmd.exe 에서 `dist\ppt2pdf_debug.exe` 입력 후 Enter.
echo 크래시 발생 시 마지막에 traceback 이 찍힙니다.
endlocal
