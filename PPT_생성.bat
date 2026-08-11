@echo off
chcp 65001 >nul
REM ============================================================
REM  Jangae weekly report - PPT generation launcher
REM  Run the Excel-update .bat first, then this one.
REM  --prev (last week's PPT path) is REQUIRED.
REM  Example:
REM     PPT.bat --date 2026-08-05 --prev "template.pptx"
REM ============================================================

set "PYEXE=%USERPROFILE%\jangae-venv\Scripts\python.exe"

if not exist "%PYEXE%" (
  echo [ERROR] venv Python not found:
  echo         %PYEXE%
  echo         See README for environment reset steps.
  pause
  exit /b 1
)

"%PYEXE%" "%~dp0make_ppt.py" %*
echo.
pause
