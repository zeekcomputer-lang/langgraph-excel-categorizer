@echo off
chcp 65001 >nul
REM ─────────────────────────────────────────────────
REM 실행 wrapper: categorizer.py 호출
REM 사용법: windows\start.bat file1.xlsx file2.xlsx ...
REM ─────────────────────────────────────────────────

cd /d "%~dp0\.."

if not exist ".venv\Scripts\python.exe" goto :need_setup

if "%OPENAI_API_KEY%"=="" goto :no_key

.venv\Scripts\python.exe tools\launch.py %*
exit /b %errorlevel%

:need_setup
echo [ERROR] .venv 가 없습니다. 먼저 windows\setup.bat 을 실행하세요.
exit /b 2

:no_key
echo [ERROR] 환경변수 OPENAI_API_KEY 가 설정되어 있지 않습니다.
echo   PowerShell:  $env:OPENAI_API_KEY = "sk-..."
echo   CMD:         set OPENAI_API_KEY=sk-...
exit /b 3
