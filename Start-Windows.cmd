@echo off
setlocal
if not exist "%~dp0.venv\Scripts\python.exe" (
  echo Run Setup-Windows.cmd first.
  if "%~1"=="" pause
  exit /b 1
)
"%~dp0.venv\Scripts\python.exe" "%~dp0main.py" %*
set "status=%ERRORLEVEL%"
if not "%status%"=="0" if "%~1"=="" pause
exit /b %status%
