@echo off
chcp 65001 >nul
REM ============================================================
REM  Jangae weekly report - Excel update launcher
REM  Runs update_excel.py using the dedicated venv Python.
REM  Usage:
REM     (no args)            base date = today
REM     --date 2026-08-05    specific base date
REM ============================================================

set "PYEXE=%USERPROFILE%\jangae-venv\Scripts\python.exe"

if not exist "%PYEXE%" (
  echo [ERROR] venv Python not found:
  echo         %PYEXE%
  echo         See README for environment reset steps.
  pause
  exit /b 1
)

"%PYEXE%" "%~dp0update_excel.py" %*
echo.
pause
