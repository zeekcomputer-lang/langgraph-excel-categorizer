@echo off
chcp 65001 >nul
REM ─────────────────────────────────────────────────
REM Windows 1회성 셋업: Python 가상환경 생성 + 의존성 설치
REM ─────────────────────────────────────────────────

cd /d "%~dp0\.."

where py >nul 2>nul
if errorlevel 1 goto :no_py

py -3 tools\setup.py
if errorlevel 1 goto :fail

echo.
echo [OK] 셋업 완료. windows\start.bat 으로 실행하세요.
exit /b 0

:no_py
echo [ERROR] Python launcher (py) 가 없습니다.
echo  https://www.python.org/downloads/ 에서 Python 3.10+ 설치 ("Add py launcher" 체크)
exit /b 1

:fail
echo [ERROR] 셋업 실패. 로그 확인 후 재실행하세요.
exit /b 1
